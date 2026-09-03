import re
import shutil
import unicodedata
from dataclasses import dataclass
from pathlib import Path

from docx import Document

PROP_RE = re.compile(
    r"PROPRIEDADE:\s*LOTE:\s*(.*?)\s*-\s*QUADRA:\s*(.*?)(?:\n|$)",
    re.IGNORECASE,
)

@dataclass(frozen=True)
class MemorialBlock:
    start: int
    end: int
    quadra: str
    lote: str
    original_index: int

def remove_accents(text: str) -> str:
    normalized = unicodedata.normalize("NFKD", text)
    return "".join(ch for ch in normalized if not unicodedata.combining(ch))

def natural_parts(text: str) -> tuple:
    clean = remove_accents(text).upper().strip()
    clean = re.sub(r"\s+", " ", clean)
    parts: list[tuple[int, object]] = []
    for piece in re.findall(r"\d+|[A-Z]+|[^A-Z\d]+", clean):
        if piece.isdigit():
            parts.append((0, int(piece)))
        elif piece.isalpha():
            parts.append((1, piece))
    return tuple(parts) if parts else ((2, clean),)

def quadra_key(quadra: str) -> tuple:
    clean = remove_accents(quadra).upper().strip()
    match = re.fullmatch(r"(\d+)\s*([A-Z]*)", clean)
    if match:
        number, suffix = match.groups()
        return (0, int(number), suffix)

    area_match = re.search(r"AREA\s+VERDE\s*(\d+)?", clean)
    if area_match:
        number = int(area_match.group(1) or 0)
        return (1, number, clean)

    return (2, natural_parts(clean))

def lote_key(lote: str) -> tuple:
    clean = remove_accents(lote).upper().strip()
    clean = re.sub(r"\s+", " ", clean)
    numbers = [int(n) for n in re.findall(r"\d+", clean)]
    first_number = numbers[0] if numbers else 10**9
    is_grouped = 1 if len(numbers) > 1 else 0

    prefix_match = re.match(r"([A-Z]+)\s*[-.]?\s*(\d+)", clean)
    if re.fullmatch(r"\d+[A-Z]?", clean):
        category = 0
        prefix = ""
    elif re.fullmatch(r"\d+\s*[-/]\s*\d+", clean):
        category = 0
        prefix = ""
    elif prefix_match:
        category = 1
        prefix = prefix_match.group(1)
    else:
        category = 2
        prefix = ""

    suffix = "".join(re.findall(r"[A-Z]+$", clean)) if numbers else ""
    return (first_number, category, is_grouped, prefix, suffix, natural_parts(clean))

def get_paragraph_text(element) -> str:
    return "".join(element.itertext())

def normalized_xml_text(element) -> str:
    text = get_paragraph_text(element).replace("\xa0", " ")
    return re.sub(r"\s+", " ", text).strip()

def find_memorial_blocks(body_elements: list) -> tuple[list, list[MemorialBlock], list]:
    starts: list[int] = []
    for index, element in enumerate(body_elements):
        if not element.tag.endswith("}p"):
            continue
        if "MEMORIAL DESCRITIVO" in normalized_xml_text(element).upper():
            starts.append(index)

    blocks: list[MemorialBlock] = []
    for original_index, start in enumerate(starts):
        end = starts[original_index + 1] if original_index + 1 < len(starts) else len(body_elements)
        text = "\n".join(get_paragraph_text(element) for element in body_elements[start:end])
        match = PROP_RE.search(text)
        if not match:
            continue
        lote = match.group(1).strip()
        quadra = match.group(2).strip()
        blocks.append(MemorialBlock(start, end, quadra, lote, original_index))

    if not blocks:
        raise ValueError("Nenhum bloco com 'PROPRIEDADE: LOTE: ... - QUADRA: ...' foi encontrado.")

    prefix = body_elements[: blocks[0].start]
    suffix = body_elements[blocks[-1].end :]
    return prefix, blocks, suffix

def sort_docx(input_path: Path, output_path: Path) -> int:
    shutil.copy2(input_path, output_path)
    document = Document(output_path)
    body = document.element.body
    body_elements = list(body)

    prefix, blocks, suffix = find_memorial_blocks(body_elements)
    block_elements = {block: body_elements[block.start : block.end] for block in blocks}
    sorted_blocks = sorted(
        blocks,
        key=lambda block: (quadra_key(block.quadra), lote_key(block.lote), block.original_index),
    )

    for element in list(body):
        body.remove(element)

    for element in prefix:
        body.append(element)
    for block in sorted_blocks:
        for element in block_elements[block]:
            body.append(element)
    for element in suffix:
        body.append(element)

    document.save(output_path)
    return len(sorted_blocks)
