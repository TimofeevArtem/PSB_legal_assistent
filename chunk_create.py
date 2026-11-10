from __future__ import annotations

import hashlib
import os
import re
import shutil
import tempfile
import zipfile

from dataclasses import dataclass
from datetime import datetime
from pathlib import Path, PurePosixPath
from typing import Optional


def _resolve_tokenizer() -> Optional[object]:
    """Создаёт локальный токенайзер в момент запуска функции, без глобального состояния."""
    try:
        import tiktoken
    except ImportError:
        return None

    try:
        return tiktoken.get_encoding("cl100k_base")
    except Exception:
        return None


def _docx_support_available() -> bool:
    """Проверяет наличие библиотеки python-docx в момент чтения документа."""
    try:
        from docx import Document
        from docx.text.paragraph import Paragraph
        from docx.table import Table
        from docx.oxml.text.paragraph import CT_P
        from docx.oxml.table import CT_Tbl
        return True
    except ImportError:
        return False

def _normalize_text(text: str) -> str:
    """Нормализует текст, приводя переводы строк и пробелы к единому формату без потери структуры абзацев."""
    if not text:
        return ""

    text = text.replace("\ufeff", "")
    text = text.replace("\r\n", "\n").replace("\r", "\n")
    text = text.replace("\xa0", " ")

    normalized_lines = []
    for line in text.split("\n"):
        leading = len(line) - len(line.lstrip(" \t"))
        content = re.sub(r"[ \t]+", " ", line.strip())
        if content:
            normalized_lines.append((" " if leading else "") + content)
        else:
            normalized_lines.append("")

    result = "\n".join(normalized_lines)
    result = re.sub(r"\n{3,}", "\n\n", result)
    return result.strip()


def _count_tokens(text: str, tokenizer: Optional[object] = None) -> int:
    """Считает число токенов в тексте с помощью локального tokenizer, либо возвращает грубую оценку."""
    if not text:
        return 0

    if tokenizer is not None:
        try:
            return len(tokenizer.encode(text))
        except Exception:
            pass

    words = re.findall(r"\S+", text)
    return max(len(words), (len(text) + 2) // 3)


def _safe_extract_path(root: Path, member_name: str) -> Path:
    """Проверяет путь внутри ZIP-архива и запрещает выход за пределы целевой директории."""
    normalized = member_name.replace("\\", "/")
    relative = PurePosixPath(normalized)

    if not normalized or relative.is_absolute():
        raise ValueError(f"Unsafe ZIP path: {member_name}")

    if any(part in {"..", "."} for part in relative.parts):
        raise ValueError(f"Unsafe ZIP path: {member_name}")

    if re.match(r"^[a-zA-Z]:", normalized):
        raise ValueError(f"Windows absolute path in ZIP: {member_name}")

    target = (root / Path(*relative.parts)).resolve()
    root_resolved = root.resolve()

    if target != root_resolved and root_resolved not in target.parents:
        raise ValueError(f"ZIP path escapes destination: {member_name}")

    return target


def _is_symlink_member(member: zipfile.ZipInfo) -> bool:
    """Определяет, является ли элемент ZIP-ссылкой, которую нельзя безопасно распаковывать."""
    unix_mode = (member.external_attr >> 16) & 0o170000
    return unix_mode == 0o120000


def extract_archive_to_directory(
    archive_path: str = "knowledge-base.zip",
    destination_dir: Optional[str] = None,
    overwrite: bool = True,
) -> list[str]:
    """Безопасно распаковывает ZIP-архив в отдельную директорию и заменяет старые данные только после успешной проверки."""
    archive = Path(archive_path).expanduser().resolve()
    destination = (
        Path(destination_dir).expanduser().resolve()
        if destination_dir
        else archive.with_suffix("")
    )
    max_archive_files = 10_000
    max_archive_size = 500 * 1024 * 1024
    max_single_file_size = 100 * 1024 * 1024
    max_compression_ratio = 1_000

    if not archive.is_file():
        raise FileNotFoundError(f"Archive not found: {archive}")

    if not zipfile.is_zipfile(archive):
        raise zipfile.BadZipFile(f"Invalid ZIP archive: {archive}")

    if destination == archive or destination in archive.parents:
        raise ValueError("Destination must not contain or replace the archive")

    destination.parent.mkdir(parents=True, exist_ok=True)

    if destination.is_symlink():
        raise ValueError("Destination directory must not be a symlink")

    members_to_extract: list[zipfile.ZipInfo] = []

    with zipfile.ZipFile(archive, "r") as zf:
        members = zf.infolist()

        if len(members) > max_archive_files:
            raise ValueError("Too many entries in ZIP archive")

        declared_total = 0
        seen_targets: set[str] = set()

        for member in members:
            if member.is_dir():
                _safe_extract_path(destination.parent / ".path-check", member.filename)
                continue

            if _is_symlink_member(member):
                continue

            if member.flag_bits & 0x1:
                raise ValueError(f"Encrypted ZIP member is not supported: {member.filename}")

            if member.file_size > max_single_file_size:
                raise ValueError(f"ZIP member exceeds size limit: {member.filename}")

            declared_total += member.file_size
            if declared_total > max_archive_size:
                raise ValueError("Declared uncompressed archive size exceeds limit")

            if member.file_size and member.compress_size == 0:
                raise ValueError(f"Suspicious compression metadata: {member.filename}")

            if member.compress_size:
                ratio = member.file_size / member.compress_size
                if ratio > max_compression_ratio:
                    raise ValueError(f"Suspicious compression ratio: {member.filename}")

            safe_target = _safe_extract_path(destination, member.filename)
            target_key = str(safe_target).casefold()
            if target_key in seen_targets:
                raise ValueError(f"Duplicate ZIP target path: {member.filename}")
            seen_targets.add(target_key)

            members_to_extract.append(member)

        temp_dir = Path(tempfile.mkdtemp(
            prefix=f".{destination.name}.staging-",
            dir=str(destination.parent),
        ))
        backup_dir = None

        try:
            actual_total = 0
            extracted: list[str] = []

            for member in members_to_extract:
                target = _safe_extract_path(temp_dir, member.filename)
                target.parent.mkdir(parents=True, exist_ok=True)

                written = 0
                with zf.open(member, "r") as source, target.open("xb") as output:
                    while True:
                        block = source.read(1024 * 1024)
                        if not block:
                            break

                        written += len(block)
                        actual_total += len(block)

                        if written > max_single_file_size:
                            raise ValueError(f"Actual file size exceeds limit: {member.filename}")

                        if actual_total > max_archive_size:
                            raise ValueError("Actual extracted archive size exceeds limit")

                        output.write(block)

                if written != member.file_size:
                    raise ValueError(f"Unexpected extracted size: {member.filename}")

                extracted.append(str(target))

            if destination.exists():
                if not overwrite:
                    raise FileExistsError(f"Destination already exists: {destination}")

                backup_dir = destination.with_name(
                    f".{destination.name}.backup-{os.getpid()}"
                )
                if backup_dir.exists():
                    shutil.rmtree(backup_dir)
                destination.rename(backup_dir)

            try:
                temp_dir.rename(destination)
            except Exception:
                if backup_dir and backup_dir.exists() and not destination.exists():
                    backup_dir.rename(destination)
                raise

            if backup_dir and backup_dir.exists():
                shutil.rmtree(backup_dir)

            return [
                str(destination / Path(p).relative_to(temp_dir))
                for p in extracted
            ]

        except Exception:
            if temp_dir.exists():
                shutil.rmtree(temp_dir, ignore_errors=True)
            raise




def _read_docx_text(file_path: Path) -> str:
    """Extract body paragraphs and tables in document order, plus headers/footers."""
    if not _docx_support_available():
        raise ImportError("Install python-docx: pip install python-docx")

    from docx import Document
    from docx.text.paragraph import Paragraph
    from docx.table import Table
    from docx.oxml.text.paragraph import CT_P
    from docx.oxml.table import CT_Tbl

    document = Document(str(file_path))
    blocks: list[str] = []

    for element in document.element.body.iterchildren():
        if isinstance(element, CT_P):
            paragraph = Paragraph(element, document)
            value = paragraph.text.strip()
            if value:
                blocks.append(value)

        elif isinstance(element, CT_Tbl):
            table = Table(element, document)
            table_lines = []

            for row in table.rows:
                cells = [
                    re.sub(r"\s+", " ", cell.text).strip()
                    for cell in row.cells
                ]
                if any(cells):
                    table_lines.append(" | ".join(cells))

            if table_lines:
                blocks.append("[ТАБЛИЦА]\n" + "\n".join(table_lines))

    header_footer_blocks = []
    for section_index, section in enumerate(document.sections, start=1):
        for kind, part in (
            ("HEADER", section.header),
            ("FOOTER", section.footer),
        ):
            values = [p.text.strip() for p in part.paragraphs if p.text.strip()]
            if values:
                header_footer_blocks.append(
                    f"[{kind} {section_index}]\n" + "\n".join(values)
                )

    if header_footer_blocks:
        blocks.extend(header_footer_blocks)

    return _normalize_text("\n\n".join(blocks))


def _read_text_file(file_path: Path) -> str:
    """Read UTF-8 first, then Windows-1251 for common Russian legacy files."""
    for encoding in ("utf-8-sig", "cp1251"):
        try:
            return file_path.read_text(encoding=encoding)
        except UnicodeDecodeError:
            continue

    raise UnicodeError(f"Unsupported text encoding: {file_path}")


def load_supported_documents_from_directory(
    folder: Path,
    supported_extensions: Optional[set[str]] = None,
    max_file_size: int = 100 * 1024 * 1024,
    skip_unreadable_documents: bool = True,
) -> list[dict]:
    """Read all supported source files from a directory and normalize them to plain text."""
    folder = folder.resolve()
    docs: list[dict] = []
    supported = supported_extensions or {".txt", ".md", ".docx"}

    for file_path in sorted(folder.rglob("*")):
        if file_path.is_symlink() or not file_path.is_file():
            continue

        try:
            resolved = file_path.resolve()
            if folder not in resolved.parents:
                continue

            if file_path.suffix.lower() not in supported:
                continue

            if file_path.stat().st_size > max_file_size:
                continue

            if file_path.suffix.lower() == ".docx":
                text = _read_docx_text(file_path)
            else:
                text = _read_text_file(file_path)

            text = _normalize_text(text)

        except Exception as error:
            message = f"Error reading {file_path}: {error}"
            if not skip_unreadable_documents:
                raise RuntimeError(message) from error
            continue

        if not text:
            continue

        title, link = _extract_document_header(file_path.name, text)

        docs.append({
            "source": str(file_path.relative_to(folder)),
            "title": title,
            "link": link,
            "text": text,
        })

    return docs

def _extract_document_header(
    file_name: str,
    text: str,
) -> tuple[str, Optional[str]]:
    """Extract a best-effort title and optional URL from the first lines."""
    lines = [line.strip() for line in text.splitlines()[:12] if line.strip()]
    if not lines:
        return Path(file_name).stem, None

    first_line = lines[0]
    link_match = re.search(r"https?://\S+|www\.\S+", first_line)

    if link_match:
        link = link_match.group(0).rstrip(".,;)")
        title = first_line[:link_match.start()].strip()
        return title or Path(file_name).stem, link

    title = first_line[:160].strip()
    return title or Path(file_name).stem, None


def _extract_document_name(file_name: str) -> str:
    """Очищает имя файла до читаемого названия документа, убирая служебные префиксы и подчеркивания."""
    stem = Path(file_name).stem
    if "__" in stem:
        stem = stem.split("__", 1)[1]
    stem = stem.replace("_", " ")
    return re.sub(r"\s+", " ", stem).strip()


def _extract_revision_date(
    file_name: str,
    text: str,
) -> Optional[str]:
    """Ищет в имени файла и тексте дату редакции по типичным юридическим паттернам."""
    header = f"{file_name}\n{text[:3000]}"
    patterns = [
        r"(?:редакци[яи]|ред\.?)\s*(?:от|:)?\s*(\d{2}\.\d{2}\.\d{4})",
        r"(?:изменени[яй]|актуализаци[яи])\s*(?:от|:)?\s*(\d{2}\.\d{2}\.\d{4})",
        r"по состоянию на\s*(\d{2}\.\d{2}\.\d{4})",
    ]

    for pattern in patterns:
        match = re.search(pattern, header, re.IGNORECASE)
        if match:
            candidate = match.group(1)
            try:
                datetime.strptime(candidate, "%d.%m.%Y")
                return candidate
            except ValueError:
                continue

    return None


def _extract_validity_status(
    file_name: str,
    text: str,
) -> str:
    """
    Deliberately conservative.

    Text heuristics cannot reliably establish legal validity. Use an external,
    authoritative verification process to set a confirmed status.
    """
    return "не проверен"


def _extract_counterparty(file_name: str) -> Optional[str]:
    """Извлекает наименование контрагента из имени договора или соглашения, если оно явно указано."""
    stem = _extract_document_name(file_name)

    if not re.search(r"договор|соглашение", stem, re.IGNORECASE):
        return None

    pattern = (
        r"(?:[A-ZА-ЯЁ]{1,8}-\d{1,4}-\d{2,4}"
        r"|[A-ZА-ЯЁ]{1,8}-\d{2,4})"
    )
    match = re.search(pattern, stem, re.IGNORECASE)

    if not match:
        return None

    after_code = stem[match.end():].strip(" _-")
    if not after_code or re.match(r"^(ред|редакция|от)\b", after_code, re.IGNORECASE):
        return None

    return after_code


def _extract_related_contract(file_name: str) -> Optional[str]:
    """Определяет номер договора, к которому относится дополнительное соглашение, если он указан в имени файла."""
    stem = _extract_document_name(file_name)

    if not re.search(r"доп|дополнительное", stem, re.IGNORECASE):
        return None

    match = re.search(
        r"к\s+договору\s*[:\-]?\s*"
        r"([A-ZА-ЯЁ]{1,8}-\d{1,4}-\d{2,4}|[A-ZА-ЯЁ]{1,8}-\d{2,4})",
        stem,
        re.IGNORECASE,
    )
    return match.group(1) if match else None


def _detect_document_type(file_name: str, text: str) -> str:
    """Определяет тип юридического документа по ключевым словам в имени и начале текста."""
    value = f"{file_name}\n{text[:1500]}".lower()

    if re.search(r"дополнительн\w*\s+соглашени", value):
        return "supplementary_agreement"
    if re.search(r"\bдоговор\b", value):
        return "contract"
    if re.search(r"\bсоглашени", value):
        return "agreement"
    if re.search(r"\b(федеральный закон|закон рф|кодекс)\b", value):
        return "law"
    if re.search(r"\b(постановление|приказ|распоряжение|указ)\b", value):
        return "regulation"
    if re.search(r"\b(политика|регламент|инструкция|положение)\b", value):
        return "policy_or_instruction"

    return "other"


def _make_content_hash(text: str) -> str:
    """Генерирует хеш содержимого, используемый для сверки и контроля дубликатов."""
    return hashlib.sha256(text.encode("utf-8")).hexdigest()


def _extract_chunk_metadata(file_name: str, text: str) -> dict:
    """Собирает базовый набор метаданных документа для последующего обогащения каждого чанка."""
    return {
        "document": _extract_document_name(file_name),
        "document_type": _detect_document_type(file_name, text),
        "counterparty": _extract_counterparty(file_name),
        "revision_date": _extract_revision_date(file_name, text),
        "validity_status": _extract_validity_status(file_name, text),
        "supplementary_agreement_for": _extract_related_contract(file_name),
        "document_content_hash": _make_content_hash(text),
    }



def _get_marker_level(marker: str) -> int:
    """Возвращает уровень вложенности юридического заголовка по его маркеру: статья, глава, раздел и т.д."""
    clean = marker.strip().rstrip(".)")

    numeric = re.fullmatch(r"\d+(?:\.\d+)*", clean)
    if numeric:
        return clean.count(".") + 1

    if re.fullmatch(r"[IVXLCDM]+", clean, re.IGNORECASE):
        return 1

    return 2


def _is_legal_heading(line: str) -> Optional[dict]:
    """Проверяет, является ли строка юридическим заголовком, и возвращает его тип, маркер и название."""
    if len(line) > 300:
        return None

    legal_heading_patterns = [
        (
            "article",
            re.compile(
                r"^\s*(?P<marker>статья\s+\d+(?:\.\d+)?)"
                r"(?:\s*[.)\-:]\s*|\s+)(?P<title>.*?)\s*$",
                re.IGNORECASE,
            ),
        ),
        (
            "chapter",
            re.compile(
                r"^\s*(?P<marker>глава\s+[IVXLCDM\d]+)"
                r"(?:\s*[.)\-:]\s*|\s+)(?P<title>.*?)\s*$",
                re.IGNORECASE,
            ),
        ),
        (
            "section",
            re.compile(
                r"^\s*(?P<marker>раздел\s+[IVXLCDM\d]+)"
                r"(?:\s*[.)\-:]\s*|\s+)(?P<title>.*?)\s*$",
                re.IGNORECASE,
            ),
        ),
        (
            "appendix",
            re.compile(
                r"^\s*(?P<marker>приложение(?:\s+№?\s*[\w\-]+)?)"
                r"(?:\s*[.)\-:]\s*|\s+)(?P<title>.*?)\s*$",
                re.IGNORECASE,
            ),
        ),
        (
            "numbered",
            re.compile(
                r"^\s*(?P<marker>"
                r"\d+(?:\.\d+)*(?:[.)])?"
                r"|[IVXLCDM]{1,8}[.)]"
                r"|[а-яёa-z][)]"
                r")\s+(?P<title>.+?)\s*$",
                re.IGNORECASE,
            ),
        ),
    ]

    for heading_type, pattern in legal_heading_patterns:
        match = pattern.match(line)
        if not match:
            continue

        marker = match.group("marker").strip()
        title = match.group("title").strip(" \t.:;-")

        if not title:
            continue

        if heading_type == "numbered":
            clean = marker.rstrip(".)")
            if re.fullmatch(r"\d+(?:\.\d+)*", clean):
                level = clean.count(".") + 1
            elif re.fullmatch(r"[IVXLCDM]+", clean, re.IGNORECASE):
                level = 1
            else:
                level = 2
        else:
            level = 1

        return {
            "type": heading_type,
            "marker": marker,
            "title": title,
            "level": level,
        }

    return None


def _numeric_parts(marker: str) -> Optional[list[str]]:
    """Извлекает числовой путь заголовка, например ['1', '2'] для '1.2' или '1.2.3'."""
    clean = marker.strip().rstrip(".)")
    if re.fullmatch(r"\d+(?:\.\d+)*", clean):
        return clean.split(".")
    return None


def _split_by_legal_sections(text: str) -> list[dict]:
    """Разбивает документ на разделы по юридическим заголовкам, сохраняя древовидную структуру и полный путь разделов."""
    sections: list[dict] = []
    stack: list[dict] = []
    current: Optional[dict] = None

    def save_current() -> None:
        """Фиксирует накопленный блок текста, если текущий раздел уже готов к сохранению."""
        nonlocal current
        if current is None:
            return

        body = "\n".join(current["body"]).strip()
        if body:
            sections.append({
                "section": current["section"],
                "section_title": current["section_title"],
                "section_path": list(current["section_path"]),
                "section_hierarchy": [dict(x) for x in current["section_hierarchy"]],
                "text": body,
            })
        current = None

    for line in text.splitlines():
        stripped = line.strip()

        if not stripped:
            if current is not None:
                current["body"].append("")
            continue

        heading = _is_legal_heading(stripped)

        if heading:
            marker = heading["marker"]
            numeric = _numeric_parts(marker)

            if numeric:
                new_stack = []
                for item in stack:
                    old_numeric = item.get("numeric_parts")
                    is_numeric_ancestor = (
                        old_numeric
                        and len(old_numeric) < len(numeric)
                        and numeric[:len(old_numeric)] == old_numeric
                    )
                    is_top_level_legal_ancestor = (
                        old_numeric is None and item["level"] == 1
                    )
                    if is_numeric_ancestor or is_top_level_legal_ancestor:
                        new_stack.append(item)
                stack = new_stack
            else:
                while stack and stack[-1]["level"] >= heading["level"]:
                    stack.pop()

            save_current()

            hierarchy = [
                {
                    "type": item["type"],
                    "marker": item["marker"],
                    "title": item["title"],
                }
                for item in stack
            ]
            hierarchy.append({
                "type": heading["type"],
                "marker": marker,
                "title": heading["title"],
            })

            current = {
                "section": marker,
                "section_title": heading["title"],
                "section_path": [item["marker"] for item in hierarchy],
                "section_hierarchy": hierarchy,
                "body": [stripped],
            }

            stack.append({
                "type": heading["type"],
                "marker": marker,
                "title": heading["title"],
                "level": heading["level"],
                "numeric_parts": numeric,
            })

        else:
            if current is None:
                current = {
                    "section": None,
                    "section_title": None,
                    "section_path": [],
                    "section_hierarchy": [],
                    "body": [],
                }
            current["body"].append(stripped)

    save_current()
    return sections



def _split_into_sentences(text: str) -> list[str]:
    """Разделяет текст на предложения, сохраняя сокращения и аббревиатуры, типичные для юридических документов."""
    text = text.strip()
    if not text:
        return []

    abbreviations = (
        "ст", "п", "подп", "ч", "г", "ул", "им", "т.д", "т.п",
        "и т.д", "и т.п", "др", "см", "рис", "табл", "№",
    )

    protected = text
    placeholders: dict[str, str] = {}

    for index, abbreviation in enumerate(abbreviations):
        escaped = re.escape(abbreviation)
        pattern = re.compile(
            rf"(?i)(?<!\w){escaped}\.(?=\s|$)"
        )

        placeholder = f"__ABBR_{index}__"
        protected, count = pattern.subn(placeholder, protected)
        if count:
            placeholders[placeholder] = abbreviation + "."

    parts = re.split(
        r"(?<=[.!?])\s+(?=[А-ЯЁA-Z0-9«\"(])",
        protected,
    )

    restored = []
    for part in parts:
        for placeholder, original in placeholders.items():
            part = part.replace(placeholder, original)
        value = part.strip()
        if value:
            restored.append(value)

    return restored


def _split_by_words(text: str, chunk_size: int) -> list[str]:
    """Разбивает текст по словам и при необходимости жёстко обрезает очень длинное слово до лимита токенов."""
    if not text.strip():
        return []

    def split_hard(value: str) -> list[str]:
        """Разбивает одно слишком длинное слово или строку на несколько фрагментов, не превышающих лимит токенов."""
        result = []
        remaining = value

        while remaining:
            if _count_tokens(remaining) <= chunk_size:
                result.append(remaining)
                break

            left, right, best = 1, len(remaining), 0
            while left <= right:
                mid = (left + right) // 2
                if _count_tokens(remaining[:mid]) <= chunk_size:
                    best = mid
                    left = mid + 1
                else:
                    right = mid - 1

            if best == 0:
                raise RuntimeError("Unable to split text within token limit")

            split_at = remaining.rfind(" ", 0, best + 1)
            if split_at > best * 0.6:
                best = split_at

            piece = remaining[:best].strip()
            if piece:
                result.append(piece)

            remaining = remaining[best:].strip()

        return result

    chunks: list[str] = []
    current: list[str] = []

    for word in text.split():
        candidate = " ".join(current + [word])

        if _count_tokens(candidate) <= chunk_size:
            current.append(word)
            continue

        if current:
            chunks.append(" ".join(current))
            current = []

        if _count_tokens(word) > chunk_size:
            chunks.extend(split_hard(word))
        else:
            current.append(word)

    if current:
        chunks.append(" ".join(current))

    return chunks


def _split_long_segment(text: str, chunk_size: int) -> list[str]:
    """Пытается разделять длинный фрагмент по предложениям, а затем по словам, чтобы сохранить смысловую структуру."""
    sentences = _split_into_sentences(text)
    if not sentences:
        return _split_by_words(text, chunk_size)

    result: list[str] = []
    current: list[str] = []

    for sentence in sentences:
        if _count_tokens(sentence) > chunk_size:
            if current:
                result.append(" ".join(current))
                current = []
            result.extend(_split_by_words(sentence, chunk_size))
            continue

        candidate = " ".join(current + [sentence])
        if _count_tokens(candidate) <= chunk_size:
            current.append(sentence)
        else:
            if current:
                result.append(" ".join(current))
            current = [sentence]

    if current:
        result.append(" ".join(current))

    return result


def _tail_within_token_limit(text: str, token_limit: int) -> str:
    """Возвращает хвост фрагмента, который умещается в лимит токенов для перекрытия соседних чанков."""
    if token_limit <= 0 or not text.strip():
        return ""

    sentences = _split_into_sentences(text)
    selected: list[str] = []

    for sentence in reversed(sentences):
        candidate = " ".join([sentence] + selected)
        if _count_tokens(candidate) <= token_limit:
            selected.insert(0, sentence)
        else:
            break

    if selected:
        return " ".join(selected)

    words = text.split()
    tail: list[str] = []

    for word in reversed(words):
        candidate = " ".join([word] + tail)
        if _count_tokens(candidate) > token_limit:
            break
        tail.insert(0, word)

    return " ".join(tail)



def _make_document_id(source: str) -> str:
    """Создаёт короткий стабильный идентификатор документа по его исходному пути."""
    return hashlib.sha256(source.encode("utf-8")).hexdigest()[:24]


def _make_parent_id(
    source: str,
    section_path: list[str],
) -> str:
    """Генерирует родительский идентификатор для группы чанков, принадлежащих одному разделу документа."""
    path = ".".join(section_path) if section_path else "general"
    value = f"{source}::{path}"
    return hashlib.sha256(value.encode("utf-8")).hexdigest()[:24]


def _make_chunk_id(source: str, index: int) -> str:
    """Создаёт уникальный идентификатор чанка в рамках исходного файла."""
    source_hash = hashlib.sha256(source.encode("utf-8")).hexdigest()[:12]
    return f"{source_hash}#{index:06d}"


def build_chunk_context(
    document_name: str,
    section_path: list[str],
    section_title: Optional[str] = None,
    section_hierarchy: Optional[list[dict]] = None,
) -> str:
    """Формирует описательный контекст чанка на основе документа, раздела и иерархии заголовков."""
    parts = []

    if document_name:
        parts.append(f"Документ: {document_name}")

    if section_hierarchy:
        hierarchy_text = " > ".join(
            f"{item.get('marker', '').strip()} {item.get('title', '').strip()}".strip()
            for item in section_hierarchy
        )
        if hierarchy_text:
            parts.append(f"Структура: {hierarchy_text}")

    elif section_path:
        path = ".".join(section_path)
        if section_title:
            parts.append(f"Раздел {path}: {section_title}")
        else:
            parts.append(f"Раздел: {path}")

    elif section_title:
        parts.append(f"Раздел: {section_title}")
    else:
        parts.append("Общее содержание документа")

    return ". ".join(parts) + "."


def _build_embedding_text(
    title: str,
    document: str,
    context: str,
    text: str,
) -> str:
    """Собирает текст для векторизации: название документа, контекст, затем сам текст чанка."""
    parts = []

    if document:
        parts.append(f"Документ: {document}")
    if title:
        parts.append(f"Название: {title}")
    if context:
        parts.append(f"Контекст: {context}")

    parts.append(text)
    return "\n".join(parts)



def _merge_small_chunks(
    chunks: list[dict],
    min_chunk_size: int,
    max_chunk_size: int,
) -> list[dict]:
    """Объединяет слишком короткие чанки с соседями в рамках одного раздела, сохраняя целостность текста."""
    return _absorb_tiny_chunks_into_neighbors(
        chunks=chunks,
        min_chunk_size=min_chunk_size,
        max_chunk_size=max_chunk_size,
    )


def _absorb_tiny_chunks_into_neighbors(
    chunks: list[dict],
    min_chunk_size: int,
    max_chunk_size: int,
) -> list[dict]:
    """Сливает слишком маленькие чанки с ближайшим совместимым соседним фрагментом, чтобы не оставлять мусорные мини-кусочки текста."""
    if not chunks:
        return []

    def compatible(left: dict, right: dict) -> bool:
        """Проверяет, можно ли безопасно объединить два чанка в одном родительском разделе."""
        left_parent = left.get("parent_id")
        right_parent = right.get("parent_id")
        if left_parent and right_parent:
            return left_parent == right_parent
        if left_parent or right_parent:
            return left_parent == right_parent
        return True

    merged: list[dict] = []
    idx = 0

    while idx < len(chunks):
        current = dict(chunks[idx])
        current["token_count"] = _count_tokens(current["text"])

        if current["token_count"] >= min_chunk_size:
            merged.append(current)
            idx += 1
            continue

        previous = merged[-1] if merged else None
        next_chunk = chunks[idx + 1] if idx + 1 < len(chunks) else None

        target = None
        target_kind = None

        if previous is not None and compatible(previous, current):
            candidate = previous["text"] + "\n\n" + current["text"]
            if _count_tokens(candidate) <= max_chunk_size:
                target = previous
                target_kind = "previous"

        if target is None and next_chunk is not None and compatible(current, next_chunk):
            candidate = current["text"] + "\n\n" + next_chunk["text"]
            if _count_tokens(candidate) <= max_chunk_size:
                target = dict(next_chunk)
                target_kind = "next"

        if target is None and previous is not None and compatible(previous, current):
            candidate = previous["text"] + "\n\n" + current["text"]
            if _count_tokens(candidate) <= max_chunk_size * 1.25:
                target = previous
                target_kind = "previous"

        if target is None and next_chunk is not None and compatible(current, next_chunk):
            candidate = current["text"] + "\n\n" + next_chunk["text"]
            if _count_tokens(candidate) <= max_chunk_size * 1.25:
                target = dict(next_chunk)
                target_kind = "next"

        if target is not None:
            if target_kind == "previous":
                previous["text"] = previous["text"] + "\n\n" + current["text"]
                previous["token_count"] = _count_tokens(previous["text"])
            else:
                target["text"] = current["text"] + "\n\n" + next_chunk["text"]
                target["token_count"] = _count_tokens(target["text"])
                chunks[idx + 1] = target
            idx += 1
            continue

        merged.append(current)
        idx += 1

    return merged


def _split_section_into_chunks(
    section_text: str,
    chunk_size: int,
    overlap: int,
    min_chunk_size: int,
) -> list[dict]:
    """Разбивает один юридический раздел на чанки, сохраняя перекрытие между соседними кусками и контекст раздела."""
    if not section_text.strip():
        return []

    overlap_budget = min(overlap, max(0, chunk_size - 1))
    content_budget = max(1, chunk_size - overlap_budget)

    raw_pieces = _split_long_segment(section_text, content_budget)
    base_chunks = [
        {"text": piece, "token_count": _count_tokens(piece)}
        for piece in raw_pieces
        if piece.strip()
    ]

    base_chunks = _absorb_tiny_chunks_into_neighbors(
        base_chunks,
        min_chunk_size=min_chunk_size,
        max_chunk_size=content_budget,
    )

    result: list[dict] = []

    for index, item in enumerate(base_chunks):
        content = item["text"]
        prefix = ""

        if index > 0 and overlap_budget > 0:
            prefix = _tail_within_token_limit(
                base_chunks[index - 1]["text"],
                overlap_budget,
            )

        final_text = f"{prefix}\n\n{content}".strip() if prefix else content

        if _count_tokens(final_text) > chunk_size:
            final_text = content

        if _count_tokens(final_text) > chunk_size:
            for piece in _split_long_segment(final_text, chunk_size):
                result.append({
                    "text": piece,
                    "token_count": _count_tokens(piece),
                })
            continue

        result.append({
            "text": final_text,
            "token_count": _count_tokens(final_text),
        })

    return result


def _split_document_into_legal_chunks(
    text: str,
    source: str = "",
    chunk_size: int = 450,
    overlap: int = 70,
    min_chunk_size: int = 50,
) -> list[dict]:
    """Разбивает весь документ на юридически структурированные чанки с привязкой к разделам и родительским идентификаторам."""
    if chunk_size <= 0:
        raise ValueError("chunk_size must be greater than zero")

    if overlap < 0 or overlap >= chunk_size:
        raise ValueError("overlap must be >= 0 and smaller than chunk_size")

    if min_chunk_size < 0:
        raise ValueError("min_chunk_size cannot be negative")

    if not text or not text.strip():
        return []

    clean = _normalize_text(text)
    sections = _split_by_legal_sections(clean)

    if not sections:
        sections = [{
            "section": None,
            "section_title": None,
            "section_path": [],
            "section_hierarchy": [],
            "text": clean,
        }]

    result: list[dict] = []

    for section in sections:
        parent_id = _make_parent_id(source, section["section_path"])

        section_chunks = _split_section_into_chunks(
            section_text=section["text"],
            chunk_size=chunk_size,
            overlap=overlap,
            min_chunk_size=min_chunk_size,
        )

        for chunk_data in section_chunks:
            token_count = _count_tokens(chunk_data["text"])

            if token_count > chunk_size:
                raise RuntimeError(
                    f"Chunk exceeds token limit: {token_count} > {chunk_size}"
                )

            result.append({
                "text": chunk_data["text"],
                "section": section["section"],
                "section_title": section["section_title"],
                "section_path": list(section["section_path"]),
                "section_hierarchy": [dict(x) for x in section["section_hierarchy"]],
                "parent_id": parent_id,
                "token_count": token_count,
            })

    return result


def _prepare_knowledge_base_from_archive(
    archive_path: Optional[str] = None,
    destination_dir: Optional[str] = None,
    chunk_size: int = 450,
    overlap: int = 70,
    min_chunk_size: int = 50,
    debug: bool = False,
    tokenizer: Optional[object] = None,
    supported_extensions: Optional[set[str]] = None,
    max_file_size: int = 100 * 1024 * 1024,
    skip_unreadable_documents: bool = True,
) -> list[dict]:
    """Запускает полный конвейер подготовки базы знаний: распаковка, чтение документов, чанкирование и обогащение метаданными."""
    if archive_path is None:
        data_dir = os.getenv("DATA_DIR")
        if not data_dir:
            raise RuntimeError("Environment variable DATA_DIR is not set")

        archive_path = str(Path(data_dir) / "knowledge-base.zip")

    archive = Path(archive_path).expanduser().resolve()
    target_dir = (
        Path(destination_dir).expanduser().resolve()
        if destination_dir
        else archive.with_suffix("")
    )

    if not archive.is_file():
        raise FileNotFoundError(f"Archive not found: {archive}")

    extract_archive_to_directory(str(archive), str(target_dir), overwrite=True)
    documents = load_supported_documents_from_directory(
        target_dir,
        supported_extensions=supported_extensions,
        max_file_size=max_file_size,
        skip_unreadable_documents=skip_unreadable_documents,
    )

    prepared: list[dict] = []

    for doc in documents:
        source = doc["source"]
        document_text = doc["text"]

        metadata = _extract_chunk_metadata(source, document_text)
        document_id = _make_document_id(source)

        structured_chunks = _split_document_into_legal_chunks(
            text=document_text,
            source=source,
            chunk_size=chunk_size,
            overlap=overlap,
            min_chunk_size=min_chunk_size,
        )

        for index, chunk_data in enumerate(structured_chunks, start=1):
            chunk_text_value = chunk_data["text"]
            section_path = chunk_data["section_path"]

            section_hint = chunk_data["section"] or "общее содержание"

            context = build_chunk_context(
                document_name=metadata["document"],
                section_path=section_path,
                section_title=chunk_data["section_title"],
                section_hierarchy=chunk_data["section_hierarchy"],
            )

            embedding_text = _build_embedding_text(
                title=doc["title"],
                document=metadata["document"],
                context=context,
                text=chunk_text_value,
            )

            item = {
                "document_id": document_id,
                "chunk_id": _make_chunk_id(source, index),
                "parent_id": chunk_data["parent_id"],
                "chunk_index": index,
                "content_hash": _make_content_hash(chunk_text_value),
                "document_content_hash": metadata["document_content_hash"],

                "source": source,
                "document": metadata["document"],
                "document_type": metadata["document_type"],
                "title": doc["title"],
                "link": doc["link"],

                "section": section_hint,
                "section_title": chunk_data["section_title"],
                "section_path": section_path,
                "section_hierarchy": chunk_data["section_hierarchy"],
                "clause_number": section_hint,

                "counterparty": metadata["counterparty"],
                "revision_date": metadata["revision_date"],
                "validity_status": metadata["validity_status"],
                "supplementary_agreement_for": metadata["supplementary_agreement_for"],

                "text": chunk_text_value,
                "context": context,
                "embedding_text": embedding_text,
                "retrieval_text": chunk_text_value,

                "token_count": _count_tokens(chunk_text_value, tokenizer=tokenizer),
                "embedding_token_count": _count_tokens(embedding_text, tokenizer=tokenizer),
                "tokenizer": "cl100k_base" if tokenizer else "fallback_estimate",
            }

            prepared.append(item)

    return prepared


def create_chunks(
    data_dir: Optional[str] = None,
    chunk_size: int = 450,
    overlap: int = 70,
    min_chunk_size: int = 50,
    debug: bool = False,
) -> list[dict]:
    """Принимает путь к папке с базой и возвращает готовые чанки для дальнейшего RAG-индекса."""
    if data_dir is None:
        data_dir = os.getenv("DATA_DIR")
        if not data_dir:
            raise RuntimeError("Environment variable DATA_DIR is not set")

    base_dir = Path(data_dir).expanduser().resolve()
    archive_path = str(base_dir / "knowledge-base.zip")
    destination_dir = str(base_dir / "knowledge-base")

    tokenizer = _resolve_tokenizer()
    supported_extensions = {".txt", ".md", ".docx"}
    max_file_size = 100 * 1024 * 1024
    skip_unreadable_documents = True

    return _prepare_knowledge_base_from_archive(
        archive_path=archive_path,
        destination_dir=destination_dir,
        chunk_size=chunk_size,
        overlap=overlap,
        min_chunk_size=min_chunk_size,
        debug=debug,
        tokenizer=tokenizer,
        supported_extensions=supported_extensions,
        max_file_size=max_file_size,
        skip_unreadable_documents=skip_unreadable_documents,
    )


if __name__ == "__main__":
    create_chunks()


    