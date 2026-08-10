#!/usr/bin/env python3
"""Shared send-run guardrails for Gmail safety."""

import json
import os
import socket
import time
from datetime import datetime


def _read_json(path):
    if not os.path.exists(path):
        return {}
    try:
        with open(path, encoding="utf-8") as f:
            data = json.load(f)
        return data if isinstance(data, dict) else {}
    except Exception:
        # Corrupt lock should not crash sends; we safely overwrite it.
        return {}


def _write_json(path, data):
    tmp = f"{path}.tmp"
    with open(tmp, "w", encoding="utf-8") as f:
        json.dump(data, f, indent=2, sort_keys=True)
    try:
        os.replace(tmp, path)
        return
    except PermissionError:
        # OneDrive-managed files on Windows can occasionally deny atomic replace.
        # Fall back to remove+replace, then to direct write as a final safety net.
        pass

    try:
        if os.path.exists(path):
            os.remove(path)
        os.replace(tmp, path)
        return
    except Exception:
        pass

    with open(path, "w", encoding="utf-8") as f:
        json.dump(data, f, indent=2, sort_keys=True)
    if os.path.exists(tmp):
        os.remove(tmp)


def enforce_send_run_cooldown(lock_path, cooldown_minutes, runner_name):
    """Block send-mode starts if the last send run began too recently."""
    if cooldown_minutes <= 0:
        return

    now_epoch = time.time()
    now_iso = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    data = _read_json(lock_path)

    last_epoch = data.get("last_send_run_started_epoch")
    last_runner = data.get("last_send_run_runner", "unknown")
    last_when = data.get("last_send_run_started_at", "unknown")

    if isinstance(last_epoch, (int, float)):
        age = now_epoch - float(last_epoch)
        window = cooldown_minutes * 60
        if age < window:
            remaining = int(window - age)
            mins = remaining // 60
            secs = remaining % 60
            raise SystemExit(
                "Cooldown guardrail: send run blocked. "
                f"Last send run started at {last_when} by {last_runner}. "
                f"Wait {mins}m {secs}s, then run again."
            )

    data.update(
        {
            "last_send_run_started_epoch": now_epoch,
            "last_send_run_started_at": now_iso,
            "last_send_run_runner": runner_name,
            "last_send_run_host": socket.gethostname(),
            "last_send_run_pid": os.getpid(),
        }
    )
    _write_json(lock_path, data)


def mark_send_run_finished(lock_path, sent_count):
    """Record completion details for observability."""
    data = _read_json(lock_path)
    data.update(
        {
            "last_send_run_finished_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "last_send_run_sent_count": int(sent_count or 0),
        }
    )
    _write_json(lock_path, data)
