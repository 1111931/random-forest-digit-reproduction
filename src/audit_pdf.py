#!/usr/bin/env python
"""Structural and layout audit for the exported manuscript PDF."""

from __future__ import annotations

import json
from pathlib import Path

import pdfplumber


ROOT = Path(__file__).resolve().parents[1]
PDF_PATH = ROOT / "paper" / "Random_Forest_Digit_Reproduction_Paper.pdf"


def main() -> None:
    expected_sections = [
        "1 引言",
        "2 相关工作",
        "3 方法",
        "4 实验",
        "5 结论",
        "参考文献",
        "附录 A",
    ]
    figure_captions = [f"图 {index}" for index in range(1, 10)]
    table_captions = [f"表 {index}" for index in range(1, 5)] + ["表 A1"]
    page_reports = []
    caption_pages: dict[str, list[int]] = {}
    full_text_parts = []
    issues = []

    with pdfplumber.open(PDF_PATH) as pdf:
        for page_number, page in enumerate(pdf.pages, start=1):
            text = page.extract_text() or ""
            words = page.extract_words()
            full_text_parts.append(text)
            off_page = [
                word
                for word in words
                if word["x0"] < 20
                or word["x1"] > page.width - 20
                or word["top"] < 20
                or word["bottom"] > page.height - 20
            ]
            if off_page:
                issues.append(
                    {
                        "page": page_number,
                        "issue": "words outside safety margin",
                        "count": len(off_page),
                    }
                )
            if len(text.strip()) < 80:
                issues.append(
                    {
                        "page": page_number,
                        "issue": "page has unexpectedly little extractable text",
                        "characters": len(text.strip()),
                    }
                )
            images = []
            for image in page.images:
                image_box = (
                    float(image["x0"]),
                    float(image["top"]),
                    float(image["x1"]),
                    float(image["bottom"]),
                )
                images.append(image_box)
                if (
                    image_box[0] < 50
                    or image_box[1] < 40
                    or image_box[2] > page.width - 50
                    or image_box[3] > page.height - 40
                ):
                    issues.append(
                        {
                            "page": page_number,
                            "issue": "image outside content region",
                            "bbox": image_box,
                        }
                    )
            lines = [line.strip() for line in text.splitlines() if line.strip()]
            page_reports.append(
                {
                    "page": page_number,
                    "characters": len(text),
                    "words": len(words),
                    "images": len(images),
                    "first_line": lines[0] if lines else "",
                    "last_line": lines[-1] if lines else "",
                }
            )
            for caption in figure_captions + table_captions:
                if caption in text:
                    caption_pages.setdefault(caption, []).append(page_number)

    full_text = "\n".join(full_text_parts)
    for section in expected_sections:
        if section not in full_text:
            issues.append({"issue": "missing section", "value": section})
    for caption in figure_captions:
        if caption not in full_text:
            issues.append({"issue": "missing figure caption", "value": caption})
        target_pages = caption_pages.get(caption, [])
        if target_pages:
            image_counts = [
                report["images"]
                for report in page_reports
                if report["page"] in target_pages
            ]
            if max(image_counts, default=0) < 1:
                issues.append(
                    {
                        "issue": "figure caption appears without an image on the same page",
                        "value": caption,
                        "page": target_pages[0],
                    }
                )
    for caption in table_captions:
        if caption not in full_text:
            issues.append({"issue": "missing table caption", "value": caption})

    required_results = [
        "86.39%",
        "95.67%",
        "98.28%",
        "89.78%",
        "97.22%",
    ]
    for value in required_results:
        if value not in full_text:
            issues.append({"issue": "missing expected result text", "value": value})

    report = {
        "pdf": str(PDF_PATH),
        "size_bytes": PDF_PATH.stat().st_size,
        "page_count": len(page_reports),
        "pages": page_reports,
        "issues": issues,
        "passed": len(issues) == 0,
    }
    output_path = ROOT / "results" / "pdf_audit.json"
    with output_path.open("w", encoding="utf-8") as handle:
        json.dump(report, handle, ensure_ascii=False, indent=2)
    print(json.dumps(report, ensure_ascii=False, indent=2))
    raise SystemExit(0 if report["passed"] else 1)


if __name__ == "__main__":
    main()
