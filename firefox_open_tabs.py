#!/usr/bin/env python3
"""Print the URLs for tabs in the current Firefox session.

By default this looks for the active Firefox profile under %APPDATA%, then reads
the newest sessionstore backup that usually reflects the current browser session.

If the session file is JSONLZ4, install lz4 first:
    pip install lz4
"""

from __future__ import annotations

import argparse
import configparser
import json
import os
from pathlib import Path
from typing import Any


def find_firefox_profiles_ini() -> Path:
    appdata = Path(os.environ.get("APPDATA", ""))
    if not appdata:
        raise SystemExit("APPDATA is not set")
    profiles_ini = appdata / "Mozilla" / "Firefox" / "profiles.ini"
    if not profiles_ini.exists():
        raise SystemExit(f"Firefox profiles.ini not found: {profiles_ini}")
    return profiles_ini


def resolve_profile_path(profiles_ini: Path) -> Path:
    parser = configparser.ConfigParser(interpolation=None)
    parser.read(profiles_ini, encoding="utf-8")

    base_dir = profiles_ini.parent
    sections = [name for name in parser.sections() if name.lower().startswith("profile")]
    if not sections:
        raise SystemExit(f"No Firefox profiles found in {profiles_ini}")

    def score(section: str) -> tuple[int, int]:
        is_default = parser.getint(section, "Default", fallback=0)
        is_relative = parser.getint(section, "IsRelative", fallback=1)
        path_value = parser.get(section, "Path", fallback="")
        path_exists = 1 if (base_dir / path_value).exists() else 0 if is_relative else 1 if Path(path_value).exists() else 0
        return (is_default, path_exists)

    chosen = max(sections, key=score)
    path_value = parser.get(chosen, "Path", fallback="")
    if not path_value:
        raise SystemExit(f"Profile Path is missing in {chosen} of {profiles_ini}")

    is_relative = parser.getint(chosen, "IsRelative", fallback=1)
    profile_path = (base_dir / path_value) if is_relative else Path(path_value)
    if not profile_path.exists():
        raise SystemExit(f"Firefox profile not found: {profile_path}")
    return profile_path


def find_session_file(profile_path: Path) -> Path:
    candidates = [
        profile_path / "sessionstore-backups" / "recovery.jsonlz4",
        profile_path / "sessionstore-backups" / "recovery.baklz4",
        profile_path / "sessionstore-backups" / "previous.jsonlz4",
        profile_path / "sessionstore.jsonlz4",
        profile_path / "sessionstore-backups" / "recovery.json",
        profile_path / "sessionstore-backups" / "previous.json",
    ]
    for candidate in candidates:
        if candidate.exists():
            return candidate
    raise SystemExit(f"No Firefox sessionstore file found under {profile_path}")


def load_session_text(session_file: Path) -> str:
    raw = session_file.read_bytes()
    if session_file.suffix.lower() in {".jsonlz4", ".baklz4"}:
        try:
            import lz4.block as lz4_block
        except ImportError as exc:
            raise SystemExit(
                "Reading .jsonlz4 session files requires the lz4 package: pip install lz4"
            ) from exc

        if raw.startswith(b"mozLz40\x00"):
            raw = raw[8:]
        return lz4_block.decompress(raw).decode("utf-8", errors="replace")

    return raw.decode("utf-8", errors="replace")


def extract_tab_urls(session_data: dict[str, Any]) -> list[dict[str, Any]]:
    tabs: list[dict[str, Any]] = []
    for window_index, window in enumerate(session_data.get("windows", []), start=1):
        for tab_index, tab in enumerate(window.get("tabs", []), start=1):
            entries = tab.get("entries", []) or []
            if not entries:
                continue
            selected_index = tab.get("index", len(entries))
            try:
                selected_index = int(selected_index)
            except (TypeError, ValueError):
                selected_index = len(entries)
            selected_index = max(1, min(selected_index, len(entries)))
            entry = entries[selected_index - 1] or {}
            tabs.append(
                {
                    "window": window_index,
                    "tab": tab_index,
                    "url": entry.get("url", ""),
                    "title": entry.get("title", ""),
                }
            )
    return tabs


def main() -> None:
    parser = argparse.ArgumentParser(description="Print URLs for open Firefox tabs in the current session.")
    parser.add_argument("--json", action="store_true", help="Print JSON instead of plain text")
    parser.add_argument("--unique", action="store_true", help="Deduplicate URLs while preserving first-seen order")
    parser.add_argument("--session-file", type=Path, help="Use a specific Firefox sessionstore file")
    args = parser.parse_args()

    if args.session_file:
        session_file = args.session_file
        if not session_file.exists():
            raise SystemExit(f"Session file not found: {session_file}")
    else:
        profile_path = resolve_profile_path(find_firefox_profiles_ini())
        session_file = find_session_file(profile_path)

    session_text = load_session_text(session_file)
    session_data = json.loads(session_text)
    tabs = extract_tab_urls(session_data)

    if args.unique:
        seen = set()
        unique_tabs = []
        for tab in tabs:
            url = tab["url"]
            if not url or url in seen:
                continue
            seen.add(url)
            unique_tabs.append(tab)
        tabs = unique_tabs

    if args.json:
        print(json.dumps(tabs, indent=2, ensure_ascii=False))
        return

    for tab in tabs:
        print(tab["url"])


if __name__ == "__main__":
    main()