import argparse
import io
import json
import re
import zipfile
from dataclasses import dataclass
from datetime import datetime, timezone
from html.parser import HTMLParser
from pathlib import Path
from typing import Iterable
from urllib.request import Request, urlopen

from docx import Document
from pypdf import PdfReader


USER_AGENT = "Mozilla/5.0"
DRIVE_FILE_RE = re.compile(r"/file/d/([A-Za-z0-9_-]+)")
DRIVE_FOLDER_RE = re.compile(r"/folders/([A-Za-z0-9_-]+)")
FOLDER_URL_RE = re.compile(r"/folders/([A-Za-z0-9_-]+)")

# Explicit overrides for names that are hard to auto-match from file names.
MANUAL_BARRIER_OVERRIDES = {
    "Precast Concrete Barrier (PCB)": {
        "fallbackLink": "__BARRIER_FOLDER__",
        "summary": "Public domain barrier entry. No product-specific TCU PDF found in the folder; open fallback folder for available documents.",
        "sourceFile": "Folder fallback",
    },
    "DB80A T150S Precast Concrete Barrier": {
        "fallbackLink": "https://drive.google.com/file/d/1Ccq2AGUOwnFThe6wyplDbI0yVjYP2Vw8/view?usp=sharing",
        "summary": "Closest available Drive TCU fallback mapped to DB80 T150S file. Verify DB80A-specific TCU file in source folder if needed.",
        "sourceFile": "211220-TCU-DB80-T150-Safety-Barrier-Temporary.pdf",
    },
}

MANUAL_END_OVERRIDES = {
    "MASH Sequential Kinking Terminal (MSKT)": {
        "fallbackLink": "https://drive.google.com/file/d/1QHeoJSZZF_volywb_mPn87HRQpxKmqoM/view?usp=sharing",
        "summary": "Product accepted MASH Sequential Kinking Terminal (MSKT). TL and use conditions are provided in the linked TCU PDF.",
        "sourceFile": "240307-TCU-MSKT-Permanent.pdf",
    },
}


@dataclass
class DriveFile:
    name: str
    url: str
    file_id: str
    path: tuple[str, ...]


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


def fetch_bytes(url: str) -> bytes:
    request = Request(url, headers={"User-Agent": USER_AGENT})
    with urlopen(request, timeout=60) as response:
        return response.read()


def fetch_text(url: str) -> str:
    return fetch_bytes(url).decode("utf-8", errors="ignore")


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


def canonical_drive_file_url(file_id: str) -> str:
    return f"https://drive.google.com/file/d/{file_id}/view?usp=sharing"


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
                file_id = file_match.group(1)
                collected.append(
                    DriveFile(
                        name=entry_title,
                        url=canonical_drive_file_url(file_id),
                        file_id=file_id,
                        path=next_path,
                    )
                )

    walk(root_id, tuple())
    return collected


def normalize_text(value: str) -> str:
    text = str(value or "").lower().replace("&", " and ")
    text = re.sub(r"\([^)]*\)", " ", text)
    text = re.sub(r"[^a-z0-9]+", " ", text)
    text = re.sub(
        r"\b(pty|ltd|limited|group|system|systems|temporary|permanent|barrier|crash|cushion|guardrail|plastic|water|filled|precast|concrete|steel|safety|end|terminal|treatment|products|civil|longitudinal|anchored|freestanding|road|roads|austroads|tcu|report)\b",
        " ",
        text,
    )
    return re.sub(r"\s+", " ", text).strip()


def token_set(value: str) -> set[str]:
    return {token for token in normalize_text(value).split() if len(token) > 1}


def aliases_for_name(name: str) -> set[str]:
    aliases = {name}
    aliases.add(re.sub(r"\s*\([^)]*\)\s*$", "", name).strip())
    aliases.add(name.replace("_", " "))
    aliases.add(re.sub(r"\bMASH TL[- ]?\d\b", "", name, flags=re.IGNORECASE).strip())
    aliases.add(re.sub(r"\bSafety Barrier\b", "", name, flags=re.IGNORECASE).strip())
    aliases.add(re.sub(r"\bCrash Cushion\b", "", name, flags=re.IGNORECASE).strip())
    aliases.add(re.sub(r"\bEnd Treatment\b", "", name, flags=re.IGNORECASE).strip())
    return {alias for alias in aliases if alias}


def score_candidate(aliases: Iterable[str], drive_file: DriveFile) -> int:
    path_text = " ".join(drive_file.path + (drive_file.name,))
    path_tokens = token_set(path_text)
    lower_name = drive_file.name.lower()
    score_best = 0

    for alias in aliases:
        alias_tokens = token_set(alias)
        if not alias_tokens:
            continue
        overlap = len(alias_tokens & path_tokens)
        score = overlap * 10
        alias_norm = normalize_text(alias)
        path_norm = normalize_text(path_text)
        if alias_norm and alias_norm in path_norm:
            score += 20
        if lower_name.endswith(".pdf"):
            score += 6
        if "tcu" in lower_name:
            score += 4
        if any(k in lower_name for k in ("manual", "datasheet", "technical", "product")):
            score += 6
        if "brochure" in lower_name:
            score -= 2
        score_best = max(score_best, score)

    return score_best


def choose_best_match(name: str, files: list[DriveFile], threshold: int = 20) -> DriveFile | None:
    aliases = aliases_for_name(name)
    ranked = sorted(
        ((score_candidate(aliases, file), file) for file in files),
        key=lambda item: item[0],
        reverse=True,
    )
    if not ranked:
        return None
    top_score, top_file = ranked[0]
    if top_score < threshold:
        return None
    return top_file


def load_database_names(barrier_data_path: Path) -> tuple[list[str], list[str]]:
    raw = barrier_data_path.read_text(encoding="utf-8-sig")
    match = re.search(r"window\.BARRIER_DATABASE\s*=\s*(\{.*\})\s*;?\s*$", raw, re.DOTALL)
    if not match:
        raise RuntimeError("Could not parse window.BARRIER_DATABASE from barrier-data.js")
    data = json.loads(match.group(1))

    barrier_rows = data["data"].get("Barrier data", [])
    end_rows = data["data"].get("End treatment", [])

    barrier_names: list[str] = []
    seen_barrier: set[str] = set()
    for row in barrier_rows[1:]:
        if not row or len(row) < 2:
            continue
        name = str(row[1] or "").strip()
        if not name or name in seen_barrier:
            continue
        seen_barrier.add(name)
        barrier_names.append(name)

    end_names: list[str] = []
    seen_end: set[str] = set()
    for row in end_rows[1:]:
        if not row or len(row) < 2:
            continue
        name = str(row[1] or "").strip()
        if not name or name in seen_end:
            continue
        seen_end.add(name)
        end_names.append(name)

    return barrier_names, end_names


def load_manufacturer_spec_links(barrier_data_path: Path) -> tuple[dict[str, str], dict[str, str]]:
    raw = barrier_data_path.read_text(encoding="utf-8-sig")
    match = re.search(r"window\.BARRIER_DATABASE\s*=\s*(\{.*\})\s*;?\s*$", raw, re.DOTALL)
    if not match:
        raise RuntimeError("Could not parse window.BARRIER_DATABASE from barrier-data.js")
    data = json.loads(match.group(1))["data"]

    barrier_links: dict[str, str] = {}
    for row in data.get("Barrier data", [])[1:]:
        if not row or len(row) < 19:
            continue
        name = str(row[1] or "").strip()
        link = str(row[18] or "").strip()
        if name and link and name not in barrier_links:
            barrier_links[name] = link

    end_links: dict[str, str] = {}
    for row in data.get("End treatment", [])[1:]:
        if not row or len(row) < 16:
            continue
        name = str(row[1] or "").strip()
        link = str(row[15] or "").strip()
        if name and link and name not in end_links:
            end_links[name] = link

    return barrier_links, end_links


def extract_drive_file_id(url: str) -> str | None:
    match = DRIVE_FILE_RE.search(str(url or ""))
    return match.group(1) if match else None


def extract_pdf_text(file_id: str) -> str:
    download_url = f"https://drive.google.com/uc?export=download&id={file_id}"
    pdf_bytes = fetch_bytes(download_url)
    reader = PdfReader(io.BytesIO(pdf_bytes))
    text_parts: list[str] = []
    # Pinning/install notes are often not on the first page; scan several pages.
    max_pages = min(8, len(reader.pages))
    for page in reader.pages[:max_pages]:
        try:
            page_text = page.extract_text() or ""
        except Exception:
            page_text = ""
        if page_text:
            text_parts.append(page_text)
        # Keep memory bounded for very large documents.
        if sum(len(part) for part in text_parts) > 120000:
            break
    return "\n".join(text_parts)


def extract_docx_text(file_bytes: bytes) -> str:
    document = Document(io.BytesIO(file_bytes))
    parts: list[str] = []
    for paragraph in document.paragraphs[:40]:
        text = (paragraph.text or "").strip()
        if text:
            parts.append(text)
    return "\n".join(parts)


def extract_drive_document_text(file_id: str) -> str:
    download_url = f"https://drive.google.com/uc?export=download&id={file_id}"
    file_bytes = fetch_bytes(download_url)

    if file_bytes.startswith(b"%PDF"):
        return extract_pdf_text(file_id)

    if file_bytes.startswith(b"PK"):
        try:
            with zipfile.ZipFile(io.BytesIO(file_bytes)) as archive:
                if "word/document.xml" in archive.namelist():
                    return extract_docx_text(file_bytes)
        except zipfile.BadZipFile:
            return ""

    return ""


def clean_text_for_summary(text: str) -> str:
    compact = re.sub(r"\s+", " ", text or "").strip()
    return compact[:8000]


def ascii_clean(text: str) -> str:
    return (
        (text or "")
        .replace("\u2013", "-")
        .replace("\u2014", "-")
        .replace("\u2018", "'")
        .replace("\u2019", "'")
        .replace("\u201c", '"')
        .replace("\u201d", '"')
    )


def find_first(pattern: str, text: str) -> str | None:
    match = re.search(pattern, text, re.IGNORECASE)
    if not match:
        return None
    return match.group(1).strip()


def extract_pinning_info(text: str) -> dict:
    """Extract installation/pinning details from TCU PDF text.

    Returns a dict with keys:
      isPinned  (bool)
      installation  (str)  e.g. "Freestanding", "Anchored at ends only", "Pinned"
      spacing   (str, optional)  e.g. "3 m", "Anchored at ends only"
      pinType   (str, optional)  e.g. "M20 x 380mm threaded rod with epoxy"
    Returns an empty dict when nothing can be determined.
    """
    cleaned = clean_text_for_summary(text)
    if not cleaned:
        return {}

    is_freestanding = bool(re.search(r"\bfree[- ]?standing\b", cleaned, re.IGNORECASE))
    has_anchor_signals = bool(re.search(
        r"\banchored\s+at\s+ends?\s+only\b|\banchor\s+spacing\b|\bpin\s+spacing\b"
        r"|\banchor\s+bolts?\b|\banchored\s+using\b|\brequires?\s+anchoring\b"
        r"|\bdriven\s+post\b|\bbaseplate\b|\bepoxy[- ]?grouted\b"
        r"|\bthreaded\s+rod\b|\bdb[- ]?pin\b|\basphalt\s+pin\b",
        cleaned, re.IGNORECASE,
    ))
    has_generic_pin_signals = bool(re.search(
        r"\bpinned\b|\bpin\s+spacing\b",
        cleaned, re.IGNORECASE,
    ))

    # Prefer explicit freestanding declarations unless there are clear anchoring signals.
    if is_freestanding and not has_anchor_signals:
        return {"isPinned": False, "installation": "Freestanding"}

    if not (has_anchor_signals or has_generic_pin_signals):
        return {}

    result: dict = {"isPinned": True}

    # Installation label
    if re.search(r"anchored\s+at\s+ends?\s+only", cleaned, re.IGNORECASE):
        result["installation"] = "Anchored at ends only"
    elif re.search(r"concrete[- ]?anchored", cleaned, re.IGNORECASE):
        result["installation"] = "Concrete Anchored"
    else:
        result["installation"] = "Pinned"

    # Pin / anchor spacing
    spacing_match = re.search(
        r"(?:pin|anchor|post)\s+spacing[^0-9]{0,20}(\d+(?:\.\d+)?)\s*m"
        r"|spacing[^0-9]{0,30}(\d+(?:\.\d+)?)\s*m"
        r"|every\s+(\d+(?:\.\d+)?)\s*m"
        r"|@\s*(\d+(?:\.\d+)?)\s*m\s*(?:c/?c|centres?|centers?)"
        r"|(\d+(?:\.\d+)?)\s*m\s+(?:c/?c|centres?|spacing)",
        cleaned, re.IGNORECASE,
    )
    if spacing_match:
        val = next((g for g in spacing_match.groups() if g), None)
        if val:
            result["spacing"] = f"{val} m"
    elif result["installation"] == "Anchored at ends only":
        result["spacing"] = "Anchored at ends only"

    # Pin / post type — most specific patterns first
    pin_patterns = [
        r"(M\d{1,2}\s*[xX\u00d7]\s*\d+\s*mm[^,.\n]{0,80})",   # M20 x 380mm threaded rod …
        r"(DB[- ]?Pin\s*\w*[^,.\n]{0,60})",                      # DB-Pin P600A
        r"((?:epoxy[- ]?grouted\s+)?threaded\s+rod[^,.\n]{0,60})",
        r"(driven\s+post[^,.\n]{0,40})",
        r"(asphalt\s+pin[^,.\n]{0,40})",
        r"(baseplate[^,.\n]{0,40})",
    ]
    for pattern in pin_patterns:
        m = re.search(pattern, cleaned, re.IGNORECASE)
        if m:
            result["pinType"] = ascii_clean(m.group(1).strip().rstrip(",.:; "))[:120]
            break

    return result


def build_summary(name: str, text: str) -> str:
    cleaned = clean_text_for_summary(text)
    if not cleaned:
        return "Summary unavailable. Use the fallback TCU file link."

    tl = find_first(r"\b((?:MASH\s*)?TL\s*-?\s*\d)\b", cleaned)
    speed = find_first(r"(?:test\s+speed|speed\s+limit|km\/?h)[^0-9]{0,20}(\d{2,3})\s*km\/?h", cleaned)
    if not speed:
        speed = find_first(r"\b(\d{2,3})\s*km\/?h\b", cleaned)

    deflection = find_first(r"(?:dynamic\s+)?deflection[^0-9]{0,20}(\d+(?:\.\d+)?)\s*m", cleaned)
    working_width = find_first(r"working\s+width[^0-9]{0,20}(\d+(?:\.\d+)?)\s*m", cleaned)
    test_length = find_first(r"(?:test\s+length|length\s+of\s+need)[^0-9]{0,20}(\d+(?:\.\d+)?)\s*m", cleaned)

    install_type = find_first(
        r"\b(freestanding|free[- ]standing|anchored\s+at\s+ends?\s+only|concrete[- ]anchored|pinned)\b",
        cleaned,
    )

    product_line = find_first(r"(Product accepted[^.]{0,180})", cleaned)
    intro = product_line or f"TCU details extracted for {name}."
    intro = ascii_clean(intro).strip().rstrip(".,;") + "."

    details: list[str] = []
    if install_type:
        details.append(f"Installation: {ascii_clean(install_type).title()}")
    if tl:
        details.append(f"TL: {re.sub(r'\\s+', ' ', tl).upper()}")
    if speed:
        details.append(f"Test speed: {speed} km/h")
    if test_length:
        details.append(f"Test length: {test_length} m")
    if deflection:
        details.append(f"Deflection: {deflection} m")
    if working_width:
        details.append(f"Working width: {working_width} m")

    summary = intro
    if details:
        summary = f"{summary} {' | '.join(details)}"
    return summary[:320]


def build_manufacturer_summary(name: str, text: str) -> str:
    cleaned = clean_text_for_summary(text)
    if not cleaned:
        return "Manufacturer specification summary unavailable. Use the manufacturer's spec link."

    cleaned = ascii_clean(cleaned)
    cleaned = re.sub(r"\b(page|issue date|revision|copyright)[^.]{0,80}\.?", " ", cleaned, flags=re.IGNORECASE)
    cleaned = re.sub(r"\s+", " ", cleaned).strip()

    sentences = [s.strip(" .;") for s in re.split(r"(?<=[.!?])\s+", cleaned) if s.strip()]
    intro = ""
    for sentence in sentences:
        if 35 <= len(sentence) <= 180:
            intro = sentence
            break
    if not intro:
        intro = f"Manufacturer spec highlights for {name}."

    width = find_first(r"(?:system\s+width|barrier\s+width|overall\s+width|width)[^0-9]{0,20}(\d+(?:\.\d+)?)\s*m", cleaned)
    length = find_first(r"(?:module\s+length|unit\s+length|length)[^0-9]{0,20}(\d+(?:\.\d+)?)\s*m", cleaned)
    weight_match = re.search(r"(?:unit\s+weight|weight|mass)[^0-9]{0,20}(\d+(?:,\d{3})*(?:\.\d+)?)\s*(kg|t|tonne|tonnes)", cleaned, re.IGNORECASE)
    install = None
    for label in ("freestanding", "anchored at ends only", "anchored", "baseplate", "driven", "epoxy"):
        if re.search(rf"\b{re.escape(label)}\b", cleaned, re.IGNORECASE):
            install = label.title()
            break

    details: list[str] = []
    if install:
        details.append(f"Installation: {install}")
    if length:
        details.append(f"Length: {length} m")
    if width:
        details.append(f"Width: {width} m")
    if weight_match:
        details.append(f"Weight: {weight_match.group(1)} {weight_match.group(2)}")

    summary = intro.rstrip(".,;:") + "."
    if details:
        summary = f"{summary} {' | '.join(details)}"
    return summary[:320]


def to_js_object_literal(data: dict) -> str:
    return json.dumps(data, ensure_ascii=True, indent=2)


def main() -> int:
    parser = argparse.ArgumentParser(description="Generate fallback TCU links and summaries from Google Drive folders")
    parser.add_argument("--barrier-folder", required=True, help="Google Drive folder URL for barrier TCU files")
    parser.add_argument("--end-folder", required=True, help="Google Drive folder URL for end treatment TCU files")
    parser.add_argument("--workspace", default=".", help="Workspace path")
    args = parser.parse_args()

    workspace = Path(args.workspace).resolve()
    barrier_data_path = workspace / "barrier-data.js"
    output_js_path = workspace / "tcu-fallback-data.js"

    barrier_names, end_names = load_database_names(barrier_data_path)
    barrier_spec_links, end_spec_links = load_manufacturer_spec_links(barrier_data_path)

    barrier_files = [f for f in walk_public_drive(args.barrier_folder) if f.name.lower().endswith(".pdf")]
    end_files = [f for f in walk_public_drive(args.end_folder) if f.name.lower().endswith(".pdf")]

    barrier_map: dict[str, dict[str, str]] = {}
    end_map: dict[str, dict[str, str]] = {}

    for name in barrier_names:
        match = choose_best_match(name, barrier_files)
        if not match:
            continue
        try:
            pdf_text = extract_pdf_text(match.file_id)
        except BaseException:
            pdf_text = ""
        barrier_map[name] = {
            "fallbackLink": match.url,
            "summary": build_summary(name, pdf_text),
            "sourceFile": match.name,
        }
        spec_file_id = extract_drive_file_id(barrier_spec_links.get(name, ""))
        if spec_file_id:
            try:
                spec_text = extract_drive_document_text(spec_file_id)
            except BaseException:
                spec_text = ""
            barrier_map[name]["manufacturerSummary"] = build_manufacturer_summary(name, spec_text)
        pinning = extract_pinning_info(pdf_text)
        if pinning:
            barrier_map[name]["pinning"] = pinning

    for name in end_names:
        match = choose_best_match(name, end_files)
        if not match:
            continue
        try:
            pdf_text = extract_pdf_text(match.file_id)
        except BaseException:
            pdf_text = ""
        end_map[name] = {
            "fallbackLink": match.url,
            "summary": build_summary(name, pdf_text),
            "sourceFile": match.name,
        }
        spec_file_id = extract_drive_file_id(end_spec_links.get(name, ""))
        if spec_file_id:
            try:
                spec_text = extract_drive_document_text(spec_file_id)
            except BaseException:
                spec_text = ""
            end_map[name]["manufacturerSummary"] = build_manufacturer_summary(name, spec_text)

    payload = {
        "generatedAt": datetime.now(timezone.utc).isoformat(),
        "sourceFolders": {
            "barriers": args.barrier_folder,
            "endTreatments": args.end_folder,
        },
        "barriers": barrier_map,
        "endTreatments": end_map,
    }

    # Apply explicit final overrides.
    for name, record in MANUAL_BARRIER_OVERRIDES.items():
        next_record = dict(record)
        if next_record.get("fallbackLink") == "__BARRIER_FOLDER__":
            next_record["fallbackLink"] = args.barrier_folder
        payload["barriers"][name] = next_record

    for name, record in MANUAL_END_OVERRIDES.items():
        payload["endTreatments"][name] = dict(record)

    output_js = "window.TCU_FALLBACK_DATA = " + to_js_object_literal(payload) + ";\n"
    output_js_path.write_text(output_js, encoding="utf-8")

    print(f"Generated {output_js_path.name}")
    print(f"Barrier entries: {len(barrier_map)} / {len(barrier_names)}")
    print(f"End treatment entries: {len(end_map)} / {len(end_names)}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
