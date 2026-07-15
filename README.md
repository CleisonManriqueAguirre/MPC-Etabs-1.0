# How to Initialize the ETABS MCP Server

This documents how to set up and launch the ETABS MCP server so Claude Desktop can drive CSI ETABS on this machine.

## 1. Layout

| Path | Purpose |
|---|---|
| `D:\MCP_etabs\etabs-mcp\server.py` | FastMCP server entry point (stdio transport) |
| `D:\MCP_etabs\etabs-mcp\etabs_client.py` | COM wrapper around ETABS (`CSI.ETABS.API.ETABSObject`) |
| `D:\MCP_etabs\etabs-mcp\engineering_checks.py` | Check registry (currently stubs — see Limitations) |
| `D:\MCP_etabs\etabs-mcp\requirements.txt` | Python dependencies for this server |
| `D:\MCP_etabs\mcp_etabs\` | Windows venv that runs the server |

## 2. Prerequisites

- CSI ETABS 23 installed at `C:\Program Files\Computers and Structures\ETABS 23` (confirmed present).
- The venv at `D:\MCP_etabs\mcp_etabs` already has all required packages installed and importable:
  `mcp`, `pywin32` (`win32com.client`), `comtypes`, `pydantic`, `python-dotenv`.

If you ever need to rebuild the venv from scratch:

```powershell
cd D:\MCP_etabs
python -m venv mcp_etabs
.\mcp_etabs\Scripts\python.exe -m pip install -r etabs-mcp\requirements.txt
.\mcp_etabs\Scripts\pywin32_postinstall.exe -install
```

The `pywin32_postinstall` step registers the COM support pywin32 needs — skip it and `win32com.client.GetActiveObject`/`Dispatch` calls can fail even though `pywin32` is installed.

## 3. Claude Desktop configuration

Your MCP server is registered in:

```
C:\Users\cleis\AppData\Roaming\Claude-3p\claude_desktop_config.json
```

```json
{
  "mcpServers": {
    "etabs": {
      "command": "D:\\MCP_etabs\\mcp_etabs\\Scripts\\python.exe",
      "args": ["D:\\MCP_etabs\\etabs-mcp\\server.py"],
      "env": {
        "ETABS_INSTALL_DIR": "C:\\Program Files\\Computers and Structures\\ETABS 23"
      }
    }
  }
}
```

**Heads up:** this machine has two separate Claude Desktop config files:
- `Claude\claude_desktop_config.json` — does **not** contain the `etabs` entry.
- `Claude-3p\claude_desktop_config.json` — contains the `etabs` entry above.

If the ETABS tools don't show up in a Claude Desktop session, confirm which config that install of Claude Desktop actually reads, and copy the `etabs` block into it if needed.

After editing either config file, fully quit and relaunch Claude Desktop (a reload isn't enough — MCP servers are spawned at app startup).

## 4. Running the server directly (manual / standalone)

You don't need Claude Desktop to start the server — it's a normal Python process that speaks MCP over stdio. Useful for checking it boots cleanly before wiring it into Claude Desktop.

```powershell
cd D:\MCP_etabs\etabs-mcp
$env:ETABS_INSTALL_DIR = "C:\Program Files\Computers and Structures\ETABS 23"
..\mcp_etabs\Scripts\python.exe server.py
```

- `server.py` calls `mcp.run(transport="stdio")`, so it blocks and waits for MCP protocol messages on stdin — running it plain like this and seeing it just hang with no errors is expected, not a bug. There's no HTTP port to open or curl.
- `ETABS_INSTALL_DIR` must be set in the environment before launch (Claude Desktop sets it via the `env` block in the config; running by hand you set it yourself as above, or via a `.env` file since `python-dotenv` is a dependency).
- To actually exercise the tools without Claude Desktop, use the MCP Inspector.

  **Do not use `mcp dev server.py` for this project.** It doesn't run your script in the `mcp_etabs` venv — it shells out to `uv run --with mcp mcp run server.py`, which spins up a throwaway `uv`-managed environment containing only the base `mcp` package. Inside that ephemeral env, `mcp run` is itself a Typer CLI command needing the `cli` extra, which `--with mcp` never installs, so it crashes immediately with `Error: typer is required. Install with 'pip install mcp[cli]'`. That plain-text error corrupts the JSON-RPC stream the Inspector expects, which shows up in the browser as `Error from MCP server: SyntaxError: Unexpected token 'E', "Error: typ"... is not valid JSON`, or more generally as `Error Connecting to MCP Inspector Proxy - Check Console logs`. Even patching that, the ephemeral env still wouldn't have `pywin32`/`comtypes`, so ETABS COM calls would fail anyway — `mcp dev`'s auto environment is simply the wrong tool here.

  Instead, launch the Inspector directly against the venv's Python, bypassing `uv` and `mcp run` entirely:

  ```powershell
  cd D:\MCP_etabs\etabs-mcp
  $env:ETABS_INSTALL_DIR = "C:\Program Files\Computers and Structures\ETABS 23"
  npx @modelcontextprotocol/inspector "D:\MCP_etabs\mcp_etabs\Scripts\python.exe" server.py
  ```

  This runs `server.py` in-process with the real venv (all of `pywin32`, `comtypes`, `pydantic`, `python-dotenv` available), with no `mcp run`/typer step involved. Open the exact URL the terminal prints (see below) to reach the Inspector UI, then call `ping`, `etabs_status`, etc. interactively.

  **"Connection Error - Check if your MCP server is running and proxy token is correct":** the terminal prints a session token and a full URL with it baked in, e.g.:

  ```
  🔑 Session token: d11cea13ad39cdadf6b0254de1b3a95e50c9a2ab0e7ba0bd61cecec76cf93623
  🚀 MCP Inspector is up and running at:
     http://localhost:6274/?MCP_PROXY_AUTH_TOKEN=d11cea13ad39cdadf6b0254de1b3a95e50c9a2ab0e7ba0bd61cecec76cf93623
  ```

  You must open that exact URL (token included). If the auto-opened browser tab landed on plain `http://localhost:6274` with no `?MCP_PROXY_AUTH_TOKEN=...`, or you refreshed/reopened the tab after a restart (the token changes every run), the proxy rejects it with this error. Copy the full URL from the terminal each time, or set `$env:DANGEROUSLY_OMIT_AUTH = "true"` before launching to disable the token check for local-only debugging.

  Also note: only one Inspector instance can bind port 6277 (proxy) / 6274 (UI) at a time — `❌ Proxy Server PORT IS IN USE at port 6277 ❌` means a previous instance is still running. On Windows, `Ctrl+C` in the terminal doesn't always kill the underlying `node` process cleanly; if the port stays busy, find and force-kill it:

  ```powershell
  Get-NetTCPConnection -LocalPort 6277,6274 -ErrorAction SilentlyContinue | Select-Object LocalPort,OwningProcess
  Stop-Process -Id <OwningProcess> -Force
  ```
- Stop the process with `Ctrl+C`.

## 5. Recommended: start ETABS first

`etabs_client.py` will auto-start ETABS via COM (`Dispatch` + `ApplicationStart`) if it isn't already running, but attaching to an **already-running** ETABS instance is more reliable:

1. Open ETABS 23 manually.
2. Open or create the model you want Claude to work with.
3. Launch/restart Claude Desktop so it spawns the MCP server.

## 6. Verifying the connection

Once Claude Desktop is running with the server registered, ask Claude to call:

- `ping` → should return `"ETABS MCP server is running."` (confirms the process launched correctly, independent of ETABS).
- `etabs_status` → returns a dict with `connected`, `mode`, `install_dir`, `prog_id`, and `error` if applicable. `connected: "true"` means it attached to a live ETABS COM object.
- `available_checks` → lists the check names currently registered (`story_drift`, `base_shear`, `modal_participation`).

## 7. Troubleshooting

| Symptom | Likely cause | Fix |
|---|---|---|
| `etabs` tools missing in Claude | Wrong config file, or Claude Desktop not restarted | Confirm which `claude_desktop_config.json` this install reads; fully restart the app |
| `pywin32 is required to connect to ETABS...` | pywin32 not installed/registered in the venv | Reinstall requirements and run `pywin32_postinstall.exe -install` |
| `etabs_status` returns `connected: false`, `mode: unavailable` | ETABS not running and nothing to attach to | Open ETABS manually, or let the client auto-start it (`allow_start=True` is the default for most calls) |
| `Unable to connect to ETABS COM server` | COM registration issue or ETABS API not licensed/installed correctly | Verify ETABS installed correctly and its COM interop is registered (reinstalling ETABS re-registers it) |
| `execute_check` returns `status: "blocked"` | Same as above — check requires a live ETABS connection | Same fix as `connected: false` above |

## 8. Current limitations (scaffold state)

- `engineering_checks.py` only confirms ETABS connectivity per check — it does not yet pull real drift, base shear, or modal mass results from `SapModel`. `execute_check` returns `status: "ready"` as a placeholder, not actual results.
- `create_simple_model` creates a blank model via `SapModel.File.NewBlank()` and can optionally save it, but has no further modeling capability yet (no grids, stories, sections, loads).
- There is no explicit `Disconnect`/`ApplicationExit` tool — the server currently only attaches or starts ETABS, it never closes it.
