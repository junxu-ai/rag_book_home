"""Utility to merge Word chapters with fresh page numbering and a Contents page.

This script demonstrates how to combine multiple Word ``.docx`` files into a
single document that contains an automatically generated *Contents* page.  Any
page numbers embedded in the source chapters are stripped so numbering can be
recalculated for the merged document.  The contents page lists each chapter
topic alongside the page on which that chapter starts.  Optionally, the merged
document can also be exported as a PDF.

Dependencies
------------
* python-docx
* docxcompose
* docx2pdf  (requires Microsoft Word or LibreOffice for conversion)
* PyPDF2

Example
-------
>>> files = ["chapter1.docx", "chapter2.docx"]
>>> topics = ["Topic 1", "Topic 2"]
>>> combine_chapters(files, topics, out_docx="book.docx", to_pdf=True)
"""
from __future__ import annotations

import os
import tempfile
from typing import List

from docx import Document
from docxcompose.composer import Composer
from docx2pdf import convert
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
from PyPDF2 import PdfReader


def _chapter_pages(files: List[str]) -> List[int]:
    """Return the number of pages for each DOCX file.

    Each file is temporarily converted to PDF so ``PyPDF2`` can determine the
    page count.  Temporary files are removed automatically.
    """
    counts = []
    temp_paths = []
    try:
        for path in files:
            fd, tmp_pdf = tempfile.mkstemp(suffix=".pdf")
            os.close(fd)
            convert(path, tmp_pdf)
            temp_paths.append(tmp_pdf)
            with open(tmp_pdf, "rb") as fh:
                reader = PdfReader(fh)
                counts.append(len(reader.pages))
    finally:
        for tmp in temp_paths:
            try:
                os.remove(tmp)
            except OSError:
                pass
    return counts


def _remove_page_numbers(doc: Document) -> None:
    """Strip existing page number fields from headers and footers."""
    for section in doc.sections:
        for hf in (section.header, section.footer):
            for p in list(hf.paragraphs):
                remove = False
                for run in p.runs:
                    text = run.text.upper()
                    elem = run._element.find('.//w:instrText', namespaces=run._element.nsmap)
                    if elem is not None:
                        text += elem.text.upper()
                    if "PAGE" in text or "NUMPAGES" in text:
                        remove = True
                        break
                if remove:
                    p._element.getparent().remove(p._element)


def _add_field_run(paragraph, field: str) -> None:
    """Insert a Word field (e.g., ``PAGE``) into ``paragraph``."""
    run = paragraph.add_run()
    fld_char = OxmlElement("w:fldChar")
    fld_char.set(qn("w:fldCharType"), "begin")
    run._r.append(fld_char)

    instr = OxmlElement("w:instrText")
    instr.set(qn("xml:space"), "preserve")
    instr.text = field
    run._r.append(instr)

    fld_char_end = OxmlElement("w:fldChar")
    fld_char_end.set(qn("w:fldCharType"), "end")
    run._r.append(fld_char_end)


def _add_page_numbers(doc: Document) -> None:
    """Insert "Page X of Y" footer numbering into ``doc``."""
    for section in doc.sections:
        footer = section.footer
        for p in list(footer.paragraphs):
            p._element.getparent().remove(p._element)
        para = footer.add_paragraph()
        para.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
        para.add_run("Page ")
        _add_field_run(para, "PAGE")
        para.add_run(" of ")
        _add_field_run(para, "NUMPAGES")


def combine_chapters(
    files: List[str],
    topics: List[str],
    *,
    out_docx: str = "combined.docx",
    to_pdf: bool = False,
) -> None:
    """Combine Word documents into a single file with a contents page.

    Parameters
    ----------
    files:
        List of ``.docx`` filenames representing each chapter.
    topics:
        List of chapter topics (same order/length as ``files``).
    out_docx:
        Output path for the merged Word document.
    to_pdf:
        If ``True`` a PDF copy of ``out_docx`` is created as well.
    """
    if len(files) != len(topics):
        raise ValueError("'files' and 'topics' must have the same length")

    # Determine the number of pages for each chapter.
    page_counts = _chapter_pages(files)

    # Compute starting page numbers (Contents occupies page 1).
    starts = []
    current = 2  # first chapter starts after the contents page
    for pages in page_counts:
        starts.append(current)
        current += pages

    # Build the master document with the contents page.
    master = Document()
    master.add_heading("Contents", level=1)
    for topic, page in zip(topics, starts):
        para = master.add_paragraph()
        para.add_run(f"{topic} ..... {page}")
    master.add_page_break()

    # Append chapters preserving original formatting but without stale page numbers.
    composer = Composer(master)
    for path in files:
        doc = Document(path)
        _remove_page_numbers(doc)
        composer.append(doc)
    composer.save(out_docx)

    # Insert new page numbers for the merged document.
    merged = Document(out_docx)
    _add_page_numbers(merged)
    merged.save(out_docx)

    if to_pdf:
        convert(out_docx, os.path.splitext(out_docx)[0] + ".pdf")


if __name__ == "__main__":
    # Example usage; replace with your own chapter files/topics.
    FILES = [r'D:\Writing\llm_rag\word_version\Ch1 Introduction.docx', 
                 r'D:\Writing\llm_rag\word_version\Ch2 LLMOps.docx', 
                 r'D:\Writing\llm_rag\word_version\Ch3 RAG.docx',
                 r'D:\Writing\llm_rag\word_version\Ch4 data process.docx', 
                 r'D:\Writing\llm_rag\word_version\Ch5 Vector.docx',
                 r'D:\Writing\llm_rag\word_version\Ch6 Query.docx',
                 r'D:\Writing\llm_rag\word_version\Ch7 retrieval.docx',
                 r'D:\Writing\llm_rag\word_version\Ch8 augumentation.docx',
                 r'D:\Writing\llm_rag\word_version\Ch9 generation.docx',
                 r'D:\Writing\llm_rag\word_version\Ch10 Evaluation.docx',
                 r'D:\Writing\llm_rag\word_version\Ch11 Serving Monitoring.docx',
                 r'D:\Writing\llm_rag\word_version\Ch12 pipeline.docx', 
                  r'D:\Writing\llm_rag\word_version\Ch13 NL2SQL.docx',
                 r'D:\Writing\llm_rag\word_version\Ch14 GraphRAG.docx',  
                r'D:\Writing\llm_rag\word_version\Ch15 Agentic RAG.docx',   
                r'D:\Writing\llm_rag\word_version\Ch16 Conclusion.docx',
                r'D:\Writing\llm_rag\word_version\Ch17 Appendix.docx',     
             ]
    TOPICS = ["Introduction",  #1
              "MLOps, LLMOps and RAGOps for Production", #2
                "RAG Challenges and Solutions", #3
                "Data Processing", #4
                "Embedding and Vector Database", # 5
                "Query Transformation and Prompt Engineering", #6
                "Retrieval Techniques", #7 
                "Augmentation and Refinement Techniques", #8 
                "Generation Techniques", #9
                "Evaluation Methodology", # 10
                "Serving and Monitoring", # 11
                "Pipeline and Orchestration", # 12
              "RAG with Database and Text2SQL", #13
              "GraphRAG", #14
              "Agentic RAG", #15
                "Conclusion", #16
                "Appendix" #17
              ]
    
    out_docx=r'D:\Writing\llm_rag\word_version\rag.docx'
    combine_chapters(FILES, TOPICS, out_docx=out_docx, to_pdf=False)
