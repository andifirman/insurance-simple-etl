from __future__ import annotations

import math
from pathlib import Path


def _pdf_escape(text: str) -> str:
    return (
        text.replace('\\', r'\\')
        .replace('(', r'\(')
        .replace(')', r'\)')
    )


def _wrap_lines(lines: list[str], max_chars: int = 95) -> list[str]:
    wrapped: list[str] = []
    for line in lines:
        if not line:
            wrapped.append("")
            continue

        s = line.rstrip("\n")
        while len(s) > max_chars:
            cut = s.rfind(" ", 0, max_chars + 1)
            if cut <= 0:
                cut = max_chars
            wrapped.append(s[:cut].rstrip())
            s = s[cut:].lstrip()
        wrapped.append(s)
    return wrapped


def build_simple_text_pdf(
    text: str,
    out_path: Path,
    title: str = "User Guide",
    page_width: int = 612,
    page_height: int = 792,
    margin: int = 54,
    font_name: str = "Helvetica",
    font_size: int = 10,
    leading: int = 14,
) -> None:
    """Write a simple multi-page PDF with left-aligned text.

    No external deps; uses standard PDF objects and built-in Helvetica.
    """

    lines = text.splitlines()
    lines = _wrap_lines(lines, max_chars=95)

    usable_height = page_height - 2 * margin
    lines_per_page = max(1, int(usable_height // leading))
    total_pages = int(math.ceil(len(lines) / lines_per_page))

    objects: list[bytes] = []

    def add_obj(payload: str) -> int:
        obj_num = len(objects) + 1
        objects.append(payload.encode("utf-8"))
        return obj_num

    # Font object
    font_obj = add_obj(
        "<< /Type /Font /Subtype /Type1 /BaseFont /{font} >>".format(font=font_name)
    )

    # Content streams + page objects
    page_objs: list[int] = []
    content_objs: list[int] = []

    for page_idx in range(total_pages):
        page_lines = lines[page_idx * lines_per_page : (page_idx + 1) * lines_per_page]

        # Build content stream
        x0 = margin
        y0 = page_height - margin - font_size

        parts: list[str] = []
        parts.append("BT")
        parts.append(f"/{'F1'} {font_size} Tf")
        parts.append(f"{x0} {y0} Td")

        # Optional title on first page
        if page_idx == 0 and title:
            parts.append(f"/{'F1'} {font_size + 4} Tf")
            parts.append(f"({_pdf_escape(title)}) Tj")
            parts.append(f"0 {-leading*2} Td")
            parts.append(f"/{'F1'} {font_size} Tf")

        for line in page_lines:
            parts.append(f"({_pdf_escape(line)}) Tj")
            parts.append(f"0 {-leading} Td")

        parts.append("ET")
        stream = "\n".join(parts).encode("utf-8")

        content_obj = add_obj(
            "<< /Length {length} >>\nstream\n{data}\nendstream".format(
                length=len(stream), data=stream.decode("utf-8")
            )
        )
        content_objs.append(content_obj)

        page_obj = add_obj(
            "<< /Type /Page /Parent {parent} 0 R /MediaBox [0 0 {w} {h}] "
            "/Resources << /Font << /F1 {font} 0 R >> >> "
            "/Contents {contents} 0 R >>".format(
                parent="{PAGES}",
                w=page_width,
                h=page_height,
                font=font_obj,
                contents=content_obj,
            )
        )
        page_objs.append(page_obj)

    # Pages object (needs kids)
    kids = " ".join([f"{p} 0 R" for p in page_objs])
    pages_obj_num = add_obj(
        "<< /Type /Pages /Kids [ {kids} ] /Count {count} >>".format(
            kids=kids, count=len(page_objs)
        )
    )

    # Replace placeholder parent ref
    fixed_objects: list[bytes] = []
    for payload in objects:
        fixed_objects.append(payload.replace(b"{PAGES}", str(pages_obj_num).encode("utf-8")))
    objects = fixed_objects

    # Catalog object
    catalog_obj = add_obj(f"<< /Type /Catalog /Pages {pages_obj_num} 0 R >>")

    # Write PDF
    out_path.parent.mkdir(parents=True, exist_ok=True)
    xref: list[int] = []

    with out_path.open("wb") as f:
        f.write(b"%PDF-1.4\n")
        for i, payload in enumerate(objects, start=1):
            xref.append(f.tell())
            f.write(f"{i} 0 obj\n".encode("ascii"))
            f.write(payload)
            f.write(b"\nendobj\n")

        xref_start = f.tell()
        f.write(f"xref\n0 {len(objects)+1}\n".encode("ascii"))
        f.write(b"0000000000 65535 f \n")
        for off in xref:
            f.write(f"{off:010d} 00000 n \n".encode("ascii"))

        f.write(
            (
                "trailer\n"
                f"<< /Size {len(objects)+1} /Root {catalog_obj} 0 R >>\n"
                "startxref\n"
                f"{xref_start}\n"
                "%%EOF\n"
            ).encode("ascii")
        )


def main() -> None:
    root = Path(__file__).resolve().parents[1]
    md_path = root / "docs" / "USER_GUIDE.md"
    out_pdf = root / "data" / "processed" / "User_Guide.pdf"

    text = md_path.read_text(encoding="utf-8")
    build_simple_text_pdf(text, out_pdf, title="Project 001 — User Guide")
    print(f"Wrote PDF: {out_pdf}")


if __name__ == "__main__":
    main()
