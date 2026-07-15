"""Local ETABS client scaffold."""

from __future__ import annotations

import os
from contextlib import suppress
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any

ETABS_PROG_ID = "CSI.ETABS.API.ETABSObject"


class EtabsConnectionError(RuntimeError):
    """Raised when the local ETABS COM server cannot be reached."""


def _load_win32_client() -> Any:
    try:
        import win32com.client as win32_client
    except ImportError as exc:  # pragma: no cover - depends on local environment
        raise EtabsConnectionError(
            "pywin32 is required to connect to ETABS. Install the requirements in the Windows venv."
        ) from exc

    return win32_client


@dataclass(slots=True)
class EtabsClient:
    """Thin wrapper around the ETABS COM object."""

    install_dir: Path | None = None
    auto_start: bool = True
    _app: Any | None = field(default=None, init=False, repr=False)
    _sap_model: Any | None = field(default=None, init=False, repr=False)

    def __post_init__(self) -> None:
        if self.install_dir is None:
            install_dir = os.getenv("ETABS_INSTALL_DIR")
            self.install_dir = Path(install_dir) if install_dir else None

    def connect(self, allow_start: bool | None = None) -> Any:
        """Attach to ETABS if possible, or start a new instance when allowed."""
        if self._app is not None:
            return self._app

        allow_start = self.auto_start if allow_start is None else allow_start
        win32_client = _load_win32_client()

        try:
            app = win32_client.GetActiveObject(ETABS_PROG_ID)
            mode = "attached"
        except Exception:
            if not allow_start:
                raise EtabsConnectionError("ETABS is not running and auto-start is disabled.")

            try:
                app = win32_client.Dispatch(ETABS_PROG_ID)
                mode = "started"
                if hasattr(app, "ApplicationStart"):
                    app.ApplicationStart()
            except Exception as exc:  # pragma: no cover - depends on local ETABS install
                raise EtabsConnectionError(f"Unable to connect to ETABS COM server: {exc}") from exc

        self._app = app
        self._sap_model = getattr(app, "SapModel", None)
        return app

    def is_connected(self) -> bool:
        """Return whether ETABS has already been attached or launched."""
        return self._app is not None

    def status(self) -> dict[str, str]:
        """Return a small status payload for Claude and debugging."""
        try:
            self.connect(allow_start=False)
            connected = "true"
            mode = "attached"
            error = ""
        except EtabsConnectionError as exc:
            connected = "false"
            mode = "unavailable"
            error = str(exc)

        payload = {
            "connected": connected,
            "mode": mode,
            "install_dir": str(self.install_dir) if self.install_dir else "not configured",
            "prog_id": ETABS_PROG_ID,
        }
        if error:
            payload["error"] = error
        return payload

    def sap_model(self) -> Any:
        """Return the connected ETABS SapModel object."""
        if self._sap_model is not None:
            return self._sap_model

        app = self.connect()
        sap_model = getattr(app, "SapModel", None)
        if sap_model is None:
            raise EtabsConnectionError("ETABS connected, but SapModel is not available on the COM object.")

        self._sap_model = sap_model
        return sap_model

    def create_simple_model(
        self,
        save_path: str | None = None,
        auto_start: bool | None = None,
    ) -> dict[str, str]:
        """Create a minimal blank ETABS model and optionally save it to disk."""
        sap_model = self.sap_model() if auto_start is None else self._connect_and_get_sap_model(auto_start)

        operation_log: list[str] = []
        with suppress(Exception):
            sap_model.InitializeNewModel()
            operation_log.append("initialized_new_model")

        file_api = getattr(sap_model, "File", None)
        if file_api is None:
            raise EtabsConnectionError("ETABS SapModel does not expose File operations.")

        new_blank = getattr(file_api, "NewBlank", None)
        if new_blank is None:
            raise EtabsConnectionError("ETABS File API does not expose NewBlank().")

        new_blank()
        operation_log.append("created_blank_model")

        resolved_save_path = self._resolve_save_path(save_path)
        if resolved_save_path is not None:
            self._save_model(file_api, resolved_save_path)
            operation_log.append("saved_model")

        return {
            "status": "created",
            "model_path": str(resolved_save_path) if resolved_save_path else "not saved",
            "connection": "ready",
            "operations": ",".join(operation_log),
        }

    def _connect_and_get_sap_model(self, allow_start: bool) -> Any:
        self.connect(allow_start=allow_start)
        return self.sap_model()

    def _resolve_save_path(self, save_path: str | None) -> Path | None:
        if not save_path:
            return None

        resolved_path = Path(save_path).expanduser().resolve()
        resolved_path.parent.mkdir(parents=True, exist_ok=True)
        return resolved_path

    def _save_model(self, file_api: Any, save_path: Path) -> None:
        save_call_error: Exception | None = None
        for candidate in (
            lambda: file_api.Save(str(save_path)),
            lambda: file_api.Save(str(save_path), 0),
        ):
            try:
                candidate()
                return
            except Exception as exc:  # pragma: no cover - depends on ETABS COM signature
                save_call_error = exc

        raise EtabsConnectionError(f"Unable to save ETABS model to {save_path}: {save_call_error}")

