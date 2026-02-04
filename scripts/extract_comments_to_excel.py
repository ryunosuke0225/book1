#!/usr/bin/env python3
"""Extract Word comments into an Excel file.

Usage:
  python scripts/extract_comments_to_excel.py input.docx output.xlsx
"""

from __future__ import annotations

import argparse
import sys
import zipfile
import xml.etree.ElementTree as ET
from dataclasses import dataclass
from datetime import datetime
from typing import Dict, List, Optional

from openpyxl import Workbook

NAMESPACES = {
    "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
}


@dataclass
class Comment:
    comment_id: str
    author: Optional[str]
    date: Optional[str]
    text: str
    commented_text: str


def _read_xml_from_docx(docx_path: str, internal_path: str) -> Optional[ET.Element]:
    try:
        with zipfile.ZipFile(docx_path) as docx_zip:
            with docx_zip.open(internal_path) as xml_file:
                return ET.fromstring(xml_file.read())
    except KeyError:
        return None


def _collect_commented_text(document_root: ET.Element) -> Dict[str, str]:
    active_comment_ids: List[str] = []
    collected: Dict[str, List[str]] = {}

    for element in document_root.iter():
        tag = element.tag
        if tag.endswith("commentRangeStart"):
            comment_id = element.attrib.get(f"{{{NAMESPACES['w']}}}id")
            if comment_id:
                active_comment_ids.append(comment_id)
                collected.setdefault(comment_id, [])
        elif tag.endswith("commentRangeEnd"):
            comment_id = element.attrib.get(f"{{{NAMESPACES['w']}}}id")
            if comment_id and comment_id in active_comment_ids:
                active_comment_ids = [cid for cid in active_comment_ids if cid != comment_id]
        elif tag.endswith("t") and element.text and active_comment_ids:
            for comment_id in active_comment_ids:
                collected.setdefault(comment_id, []).append(element.text)

    return {comment_id: "".join(texts).strip() for comment_id, texts in collected.items()}


def _extract_comments(docx_path: str) -> List[Comment]:
    comments_root = _read_xml_from_docx(docx_path, "word/comments.xml")
    if comments_root is None:
        return []

    document_root = _read_xml_from_docx(docx_path, "word/document.xml")
    commented_text_map: Dict[str, str] = {}
    if document_root is not None:
        commented_text_map = _collect_commented_text(document_root)

    comments: List[Comment] = []
    for comment in comments_root.findall("w:comment", NAMESPACES):
        comment_id = comment.attrib.get(f"{{{NAMESPACES['w']}}}id", "")
        author = comment.attrib.get(f"{{{NAMESPACES['w']}}}author")
        date_value = comment.attrib.get(f"{{{NAMESPACES['w']}}}date")
        formatted_date = None
        if date_value:
            try:
                formatted_date = datetime.fromisoformat(date_value.replace("Z", "+00:00")).isoformat()
            except ValueError:
                formatted_date = date_value

        text_runs = [node.text or "" for node in comment.findall(".//w:t", NAMESPACES)]
        text = "".join(text_runs).strip()
        commented_text = commented_text_map.get(comment_id, "")

        comments.append(
            Comment(
                comment_id=comment_id,
                author=author,
                date=formatted_date,
                text=text,
                commented_text=commented_text,
            )
        )

    return comments


def _write_to_excel(comments: List[Comment], output_path: str) -> None:
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "Comments"

    headers = ["Comment ID", "Author", "Date", "Comment", "Commented Text"]
    sheet.append(headers)

    for comment in comments:
        sheet.append(
            [
                comment.comment_id,
                comment.author or "",
                comment.date or "",
                comment.text,
                comment.commented_text,
            ]
        )

    workbook.save(output_path)


def main() -> int:
    parser = argparse.ArgumentParser(description="Extract Word comments into an Excel file.")
    parser.add_argument("input_docx", help="Path to the input .docx file")
    parser.add_argument("output_xlsx", help="Path to the output .xlsx file")
    args = parser.parse_args()

    comments = _extract_comments(args.input_docx)
    if not comments:
        print("No comments found or comments.xml missing.")

    _write_to_excel(comments, args.output_xlsx)
    print(f"Saved {len(comments)} comment(s) to {args.output_xlsx}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
