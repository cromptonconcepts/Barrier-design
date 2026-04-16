import argparse
import json
import re
import sys
from dataclasses import dataclass
from html.parser import HTMLParser
from pathlib import Path
from typing import Iterable
from urllib.error import HTTPError, URLError
from urllib.request import Request, urlopen


USER_AGENT = "Mozilla/5.0"
DRIVE_FILE_RE = re.compile(r"/file/d/([A-Za-z0-9_-]+)")
DRIVE_FOLDER_RE = re.compile(r"/folders/([A-Za-z0-9_-]+)")
FOLDER_URL_RE = re.compile(r"/folders/([A-Za-z0-9_-]+)")
JSON_INDENT = 4


@dataclass
class DriveFile:
    name: str
    url: str
    path: tuple[str, ...]


@dataclass
class MatchResult:
    dataset: str
    item_name: str
    old_url: str | None
    new_url: str
    score: int
    folder_path: str


class EmbeddedFolderParser(HTMLParser):
    def __init__(self) -> None:
        super().__init__()
        self.entries: list[dict[str, str]] = []
        self._current_href: str | None = None
        self._capture_title = False
        self._title_parts: list[str] = []

    def handle_starttag(self, tag: str, attrs: list[tuple[str, str | None]]) -> None:
        attrs_dict = dict(attrs)
        if tag == "a":
            href = attrs_dict.get("href")
            if href and (
                "drive.google.com/drive/folders/" in href
                or "drive.google.com/file/d/" in href
            ):
                self._current_href = href
                self._title_parts = []
        elif tag == "div" and self._current_href:
            classes = attrs_dict.get("class") or ""
            if "flip-entry-title" in classes:
                self._capture_title = True

    def handle_data(self, data: str) -> None:
        if self._capture_title:
            self._title_parts.append(data)

    def handle_endtag(self, tag: str) -> None:
        if tag == "div" and self._capture_title:
            self._capture_title = False
            title = " ".join(part.strip() for part in self._title_parts if part.strip())
            if self._current_href and title:
                self.entries.append({"href": self._current_href, "title": title})
                self._current_href = None
                self._title_parts = []
        elif tag == "a" and self._current_href and not self._capture_title:
            self._current_href = None
            self._title_parts = []


def fetch_text(url: str) -> str:
    request = Request(url, headers={"User-Agent": USER_AGENT})
    with urlopen(request, timeout=30) as response:
        return response.read().decode("utf-8", errors="ignore")


def extract_folder_id(url: str) -> str:
    match = FOLDER_URL_RE.search(url)
    if not match:
        raise ValueError(f"Could not extract Google Drive folder ID from: {url}")
    return match.group(1)


def embedded_folder_url(folder_id: str) -> str:
    return f"https://drive.google.com/embeddedfolderview?id={folder_id}#list"


def parse_embedded_folder(html: str) -> tuple[str, list[dict[str, str]]]:
    title_match = re.search(r"<title>(.*?)</title>", html, re.IGNORECASE | re.DOTALL)
    title = (title_match.group(1).strip() if title_match else "Untitled Folder")
    parser = EmbeddedFolderParser()
    parser.feed(html)
    unique_entries: list[dict[str, str]] = []
    seen: set[tuple[str, str]] = set()
    for entry in parser.entries:
        key = (entry["href"], entry["title"])
        if key in seen:
            continue
        seen.add(key)
        unique_entries.append(entry)
    return title, unique_entries


def canonical_drive_file_url(url: str) -> str:
    match = DRIVE_FILE_RE.search(url)
    if not match:
        return url
    return f"https://drive.google.com/file/d/{match.group(1)}/view?usp=sharing"


def walk_public_drive(folder_url: str) -> list[DriveFile]:
    root_id = extract_folder_id(folder_url)
    collected: list[DriveFile] = []
    visited: set[str] = set()

    def walk(folder_id: str, path_parts: tuple[str, ...]) -> None:
        if folder_id in visited:
            return
        visited.add(folder_id)
        html = fetch_text(embedded_folder_url(folder_id))
        title, entries = parse_embedded_folder(html)
        next_path = path_parts + (title,)
        for entry in entries:
            href = entry["href"]
            entry_title = entry["title"].strip()
            folder_match = DRIVE_FOLDER_RE.search(href)
            if folder_match:
                walk(folder_match.group(1), next_path)
                continue
            file_match = DRIVE_FILE_RE.search(href)
            if file_match:
                collected.append(
                    DriveFile(
                        name=entry_title,
                        url=canonical_drive_file_url(href),
                        path=next_path,
                    )
                )

    walk(root_id, tuple())
    return collected


def normalize_text(value: str) -> str:
    value = value.lower().replace("&", " and ")
    value = re.sub(r"[^a-z0-9]+", " ", value)
    value = re.sub(r"\b(pty|ltd|limited|group|system|systems|temporary|permanent|barrier|crash|cushion|guardrail|plastic|water|filled|precast|concrete|steel|safety|end|terminal|treatment|products|civil|longitudinal|anchored|freestanding|road|roads)\b", " ", value)
    return re.sub(r"\s+", " ", value).strip()


def token_set(value: str) -> set[str]:
    normalized = normalize_text(value)
    return {token for token in normalized.split() if len(token) > 1}


def aliases_for_barrier(name: str, supplier: str | None = None, family: str | None = None) -> set[str]:
    aliases = {name}
    short = re.sub(r"\s*[-–]\s*temporary\s*$", "", name, flags=re.IGNORECASE).strip()
    aliases.add(short)
    aliases.add(re.sub(r"\bprecast concrete barrier\b", "", name, flags=re.IGNORECASE).strip())
    aliases.add(re.sub(r"\bsafety barrier\b", "", name, flags=re.IGNORECASE).strip())
    aliases.add(re.sub(r"\bsteel and concrete longitudinal barrier\b", "", name, flags=re.IGNORECASE).strip())
    if family:
        aliases.add(family)
    if supplier:
        aliases.add(f"{supplier} {name}")
    return {alias for alias in aliases if alias}


def aliases_for_end_treatment(name: str, supplier: str | None = None, family: str | None = None) -> set[str]:
    aliases = {name}
    aliases.add(re.sub(r"\bcrash cushion\b", "", name, flags=re.IGNORECASE).strip())
    aliases.add(re.sub(r"\bguardrail end terminal\b", "", name, flags=re.IGNORECASE).strip())
    aliases.add(re.sub(r"\bmash sequential kinking terminal\b", "MSKT", name, flags=re.IGNORECASE).strip())
    aliases.add(re.sub(r"\bquadguard\b", "QUADGUARD", name, flags=re.IGNORECASE).strip())
    if family:
        aliases.add(family)
    if supplier:
        aliases.add(f"{supplier} {name}")
    return {alias for alias in aliases if alias}


def score_candidate(aliases: Iterable[str], drive_file: DriveFile) -> int:
    path_text = " ".join(drive_file.path + (drive_file.name,))
    path_tokens = token_set(path_text)
    best = 0
    normalized_path = normalize_text(path_text)
    for alias in aliases:
        normalized_alias = normalize_text(alias)
        alias_tokens = {token for token in normalized_alias.split() if token}
        if not alias_tokens:
            continue
        overlap = len(alias_tokens & path_tokens)
        score = overlap * 10
        if normalized_alias and normalized_alias in normalized_path:
            score += 25
        if drive_file.name.lower().endswith(".pdf"):
            score += 5
        if len(drive_file.path) > 1:
            score += min(len(drive_file.path), 4)
        lower_name = drive_file.name.lower()
        if any(keyword in lower_name for keyword in ("manual", "datasheet", "data-sheet", "data_sheet", "tech", "technical", "product")):
            score += 8
        if "brochure" in lower_name:
            score -= 2
        if "tcu" in lower_name:
            score -= 4
        if re.match(r"^\d+m[_ -]", lower_name):
            score -= 3
        if lower_name.endswith((".jpg", ".jpeg", ".png", ".webp")):
            score -= 10
        best = max(best, score)
    return best


def candidate_preference(drive_file: DriveFile) -> tuple[int, str]:
    lower_name = drive_file.name.lower()
    preference = 0
    if "installation" in lower_name and "manual" in lower_name:
        preference += 6
    elif "product" in lower_name and "manual" in lower_name:
        preference += 5
    elif "manual" in lower_name:
        preference += 4
    elif any(keyword in lower_name for keyword in ("datasheet", "data-sheet", "data_sheet", "tech", "technical")):
        preference += 3
    elif "product" in lower_name:
        preference += 2
    elif "flyer" in lower_name:
        preference += 1
    if "brochure" in lower_name:
        preference -= 1
    if "tcu" in lower_name:
        preference -= 2
    return preference, lower_name


def best_match(aliases: Iterable[str], files: list[DriveFile], threshold: int = 20) -> DriveFile | None:
    ranked = sorted(
        ((score_candidate(aliases, drive_file), candidate_preference(drive_file), drive_file) for drive_file in files),
        key=lambda item: (item[0], item[1][0], item[1][1]),
        reverse=True,
    )
    if not ranked:
        return None
    top_score, _, top_file = ranked[0]
    if top_score < threshold:
        return None
    return top_file


def load_barrier_database(path: Path) -> dict:
    raw = path.read_text(encoding="utf-8-sig")
    match = re.search(r"window\.BARRIER_DATABASE\s*=\s*(\{.*\})\s*;?\s*$", raw, re.DOTALL)
    if not match:
        raise ValueError("Could not parse window.BARRIER_DATABASE from barrier-data.js")
    payload = match.group(1).strip()
    if payload.endswith(";"):
        payload = payload[:-1]
    return json.loads(payload)


def save_barrier_database(path: Path, database: dict) -> None:
    content = "window.BARRIER_DATABASE = " + json.dumps(database, indent=JSON_INDENT, ensure_ascii=True) + ";\n"
    path.write_text(content, encoding="utf-8")


def extract_end_treatments_from_html(path: Path) -> tuple[str, list[dict], str]:
    raw = path.read_text(encoding="utf-8")
    match = re.search(r"(const END_TREATMENTS = )(\[.*?\])(\s*;\s*let selectedEndTreatment = null;)", raw, re.DOTALL)
    if not match:
        raise ValueError("Could not locate END_TREATMENTS array in index.html")
    prefix, payload, suffix = match.groups()
    return raw, json.loads(payload), prefix + "__PAYLOAD__" + suffix


def save_end_treatments_to_html(path: Path, original_raw: str, template: str, end_treatments: list[dict]) -> None:
    replacement = template.replace("__PAYLOAD__", json.dumps(end_treatments, indent=JSON_INDENT, ensure_ascii=True))
    updated = re.sub(r"const END_TREATMENTS = \[.*?\]\s*;\s*let selectedEndTreatment = null;", replacement, original_raw, count=1, flags=re.DOTALL)
    path.write_text(updated, encoding="utf-8")


def update_barrier_rows(database: dict, barrier_files: list[DriveFile]) -> list[MatchResult]:
    rows = database["data"]["Barrier data"]
    changes: list[MatchResult] = []
    for index, row in enumerate(rows[1:], start=1):
        if len(row) < 19:
            continue
        family = (row[0] or "").strip() if row[0] else None
        name = (row[1] or "").strip() if row[1] else ""
        supplier = (row[15] or "").strip() if len(row) > 15 and row[15] else None
        if not name:
            continue
        aliases = aliases_for_barrier(name, supplier, family)
        match = best_match(aliases, barrier_files)
        if not match:
            continue
        old_url = row[18] if len(row) > 18 else None
        if old_url == match.url:
            continue
        row[18] = match.url
        changes.append(
            MatchResult(
                dataset="barrier",
                item_name=name,
                old_url=old_url,
                new_url=match.url,
                score=score_candidate(aliases, match),
                folder_path=" / ".join(match.path),
            )
        )
    return changes


def update_end_treatment_rows(database: dict, end_files: list[DriveFile]) -> list[MatchResult]:
    rows = database["data"]["End treatment"]
    changes: list[MatchResult] = []
    for index, row in enumerate(rows[1:], start=1):
        if len(row) < 16:
            continue
        family = (row[0] or "").strip() if row[0] else None
        name = (row[1] or "").strip() if row[1] else ""
        supplier = (row[13] or "").strip() if len(row) > 13 and row[13] else None
        if not name:
            continue
        aliases = aliases_for_end_treatment(name, supplier, family)
        match = best_match(aliases, end_files)
        if not match:
            continue
        old_url = row[15] if len(row) > 15 else None
        if old_url == match.url:
            continue
        row[15] = match.url
        changes.append(
            MatchResult(
                dataset="end-treatment-sheet",
                item_name=name,
                old_url=old_url,
                new_url=match.url,
                score=score_candidate(aliases, match),
                folder_path=" / ".join(match.path),
            )
        )
    return changes


def update_index_end_treatments(end_treatments: list[dict], end_files: list[DriveFile]) -> list[MatchResult]:
    changes: list[MatchResult] = []
    for record in end_treatments:
        name = (record.get("name") or "").strip()
        family = (record.get("family") or record.get("shortName") or None)
        supplier = (record.get("supplier") or "").strip() or None
        if not name:
            continue
        old_url = record.get("manufLink")
        if old_url and isinstance(old_url, str) and "drive.google.com/file/d/" in old_url and "cromptonconcepts-my.sharepoint.com" not in old_url:
            continue
        aliases = aliases_for_end_treatment(name, supplier, family)
        match = best_match(aliases, end_files)
        if not match:
            continue
        if old_url == match.url:
            continue
        record["manufLink"] = match.url
        changes.append(
            MatchResult(
                dataset="index-end-treatment",
                item_name=name,
                old_url=old_url,
                new_url=match.url,
                score=score_candidate(aliases, match),
                folder_path=" / ".join(match.path),
            )
        )
    return changes


def print_report(changes: list[MatchResult], label: str) -> None:
    print(f"\n{label}: {len(changes)} updates")
    for change in sorted(changes, key=lambda item: (item.dataset, item.item_name.lower())):
        print(f"- [{change.dataset}] {change.item_name}")
        print(f"  score: {change.score}")
        print(f"  folder: {change.folder_path}")
        print(f"  old: {change.old_url or '(blank)'}")
        print(f"  new: {change.new_url}")


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Update manufacturer spec links from public Google Drive folders.")
    parser.add_argument("--barrier-folder", required=True, help="Google Drive folder URL for barriers")
    parser.add_argument("--end-folder", required=True, help="Google Drive folder URL for end treatments")
    parser.add_argument("--workspace", default=".", help="Workspace folder containing barrier-data.js and index.html")
    parser.add_argument("--write", action="store_true", help="Persist changes to files. Without this flag, runs as a dry run.")
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    workspace = Path(args.workspace).resolve()
    barrier_data_path = workspace / "barrier-data.js"
    index_path = workspace / "index.html"

    try:
        barrier_files = walk_public_drive(args.barrier_folder)
        end_files = walk_public_drive(args.end_folder)
    except (HTTPError, URLError, ValueError) as exc:
        print(f"Failed to read Google Drive folders: {exc}", file=sys.stderr)
        return 1

    database = load_barrier_database(barrier_data_path)
    original_html, end_treatments, html_template = extract_end_treatments_from_html(index_path)

    barrier_changes = update_barrier_rows(database, barrier_files)
    end_sheet_changes = update_end_treatment_rows(database, end_files)
    index_changes = update_index_end_treatments(end_treatments, end_files)
    all_changes = barrier_changes + end_sheet_changes + index_changes

    print(f"Barrier files discovered: {len(barrier_files)}")
    print(f"End-treatment files discovered: {len(end_files)}")
    print_report(all_changes, "Planned changes")

    if args.write and all_changes:
        save_barrier_database(barrier_data_path, database)
        save_end_treatments_to_html(index_path, original_html, html_template, end_treatments)
        print("\nFiles updated.")
    elif args.write:
        print("\nNo changes written because no updates were identified.")
    else:
        print("\nDry run only. Re-run with --write to persist changes.")

    return 0


if __name__ == "__main__":
    raise SystemExit(main())