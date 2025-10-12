# -*- coding: utf-8 -*-
"""
Merges multiple .docx files into a single document with a Table of Contents,
preserving the original formatting of the chapter content.

This script combines specified .docx files into a single Word document. It is
designed to maintain the formatting of the source content (text styles, tables,
images) while creating a unified final document.

Key Features:
- Preserves original content formatting: Fonts, colors, paragraph styles,
  tables, and images from the source files are kept intact.
- Starts each chapter on a new page: Uses a robust 'Next Page' section break
  to ensure clean separation between chapters.
- Normalizes page layout: To ensure a consistent look and continuous page
  numbering for the TOC, headers, footers, and margins from source files are
  not used. The final document has its own unified page layout.
- Flexible chapter definition: Merge files from a folder, a specific list of
  files, or a list of files with custom titles.
- Optional PDF conversion with an updated Table of Contents (Windows-only).

Usage:
1. Place your .docx files in a known location (e.g., 'source_chapters' folder).
2. Configure the `INPUT_MODE` and other settings in the main execution block
   at the bottom of the script.
3. Run the script.
"""
import os
import sys
from docx import Document
from docx.enum.section import WD_SECTION
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

# --- Helper Function for Table of Contents ---

def add_table_of_contents(doc):
    """
    Adds a table of contents field to the document.
    Word must be used to update this field to generate the actual TOC.
    """
    doc.add_heading('Contents', level=1)
    paragraph = doc.add_paragraph()
    run = paragraph.add_run()
    fldChar_begin = OxmlElement('w:fldChar')
    fldChar_begin.set(qn('w:fldCharType'), 'begin')
    instrText = OxmlElement('w:instrText')
    instrText.set(qn('xml:space'), 'preserve')
    instrText.text = 'TOC \\o "1-3" \\h \\z \\u'
    fldChar_separate = OxmlElement('w:fldChar')
    fldChar_separate.set(qn('w:fldCharType'), 'separate')
    fldChar_end = OxmlElement('w:fldChar')
    fldChar_end.set(qn('w:fldCharType'), 'end')
    run._r.append(fldChar_begin)
    run._r.append(instrText)
    run._r.append(fldChar_separate)
    run = paragraph.add_run('Right-click and select "Update Field" to generate the table of contents.')
    run.font.italic = True
    run = paragraph.add_run()
    run._r.append(fldChar_end)

# --- Main Merging Function ---

def merge_documents(chapters, final_doc_path):
    """
    Merges a list of documents into a single document, preserving content
    formatting and starting each chapter on a new page.
    """
    if not chapters:
        print("The list of chapters to merge is empty. Aborting.")
        return None

    # Create the master document using the first chapter's document as a template
    # This helps in preserving the initial section's properties if needed.
    first_doc_path, _ = chapters[0]
    if not os.path.exists(first_doc_path):
        print(f"Error: The first document '{first_doc_path}' was not found. Aborting.")
        return None
        
    master_document = Document()
    add_table_of_contents(master_document)
    
    print("Starting the merge process...")
    
    for i, (filepath, chapter_title) in enumerate(chapters):
        print(f"  -> Merging '{filepath}' as '{chapter_title}'")
        
        if not os.path.exists(filepath):
            print(f"    [!] WARNING: File not found, skipping: {filepath}")
            continue

        # Add a "Next Page" section break before each new chapter (except the first)
        if i > 0:
            master_document.add_section(WD_SECTION.NEW_PAGE)

        # Add the chapter title as a Heading 1 for the TOC
        master_document.add_heading(chapter_title, level=1)
        
        # Open the source document
        source_doc = Document(filepath)
        
        # Append content from source to master.
        # This loop copies paragraphs and tables, preserving their formatting.
        for element in source_doc.element.body:
            master_document.element.body.append(element)

    print("\nMerge complete. Saving the final Word document...")
    master_document.save(final_doc_path)
    print(f"✅ Successfully saved to '{final_doc_path}'")
    
    return final_doc_path


# --- PDF Conversion Function (Windows Only) ---

def convert_to_pdf(docx_path, pdf_path):
    """
    Converts a .docx file to .pdf using Microsoft Word.
    This function is Windows-only and requires Word to be installed.
    """
    if sys.platform != 'win32':
        print("\nℹ️ PDF conversion is only available on Windows.")
        return

    try:
        import comtypes.client
    except ImportError:
        print("\nCould not convert to PDF. Please install 'comtypes' and 'pywin32' libraries.")
        return
        
    word, doc = None, None
    try:
        print("\nStarting PDF conversion...")
        word = comtypes.client.CreateObject('Word.Application')
        word.Visible = False
        abs_docx_path = os.path.abspath(docx_path)
        abs_pdf_path = os.path.abspath(pdf_path)
        doc = word.Documents.Open(abs_docx_path)
        print("  -> Updating Table of Contents...")
        doc.Fields.Update()
        print(f"  -> Saving PDF to '{pdf_path}'...")
        doc.SaveAs(abs_pdf_path, FileFormat=17) # 17 is the wdFormatPDF value
        print("✅ PDF conversion successful.")
    except Exception as e:
        print(f"\n❌ An error occurred during PDF conversion: {e}")
        print("Please ensure Microsoft Word is installed and not running with dialogs open.")
    finally:
        if doc: doc.Close(SaveChanges=False)
        if word: word.Quit()


# --- Main Execution Block ---

if __name__ == '__main__':
    # ========================== CONFIGURATION ==========================
    
    # --- Step 1: Choose your Input Mode ---
    # Options: 'FOLDER', 'FILE_LIST', 'CUSTOM'
    INPUT_MODE = 'CUSTOM'

    # --- Step 2: Set General Parameters ---
    SOURCE_FOLDER = 'source_chapters'
    OUTPUT_DOCX_FILENAME = 'merged_document_formatted.docx'
    OUTPUT_PDF_FILENAME = 'merged_document_formatted.pdf'
    CREATE_PDF = False # Set to True for Windows with MS Word

    # --- Step 3: Configure the selected Input Mode ---
    
    # Config for 'FILE_LIST' mode
    files_for_list_mode = [
        '01_Introduction.docx',
        '02_System_Architecture.docx',
        '03_Conclusion.docx',
    ]

    # Config for 'CUSTOM' mode
    files_for_custom_mode = [
                r'D:\Writing\llm_rag\word_version\Ch1 Introduction.docx', 
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
    titles_for_custom_mode = [
        "Introduction",  #1
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


    # ======================= END OF CONFIGURATION =======================
    
    if INPUT_MODE == 'FOLDER':
        if not os.path.exists(SOURCE_FOLDER):
            os.makedirs(SOURCE_FOLDER)
            print(f"Created folder '{SOURCE_FOLDER}'. Please add your .docx files and run again.")
            sys.exit()
        
    chapters_to_merge = []
    
    if INPUT_MODE == 'FOLDER':
        print(f"Mode: FOLDER. Reading all .docx files from '{SOURCE_FOLDER}'...")
        source_files = sorted([f for f in os.listdir(SOURCE_FOLDER) if f.endswith('.docx')])
        if not source_files:
             print(f"No .docx files found in '{SOURCE_FOLDER}'.")
             sys.exit()
        for filename in source_files:
            chapters_to_merge.append((os.path.join(SOURCE_FOLDER, filename), os.path.splitext(filename)[0]))
            
    elif INPUT_MODE == 'FILE_LIST':
        print(f"Mode: FILE_LIST. Using specified file order...")
        for filename in files_for_list_mode:
            chapters_to_merge.append((os.path.join(SOURCE_FOLDER, filename), os.path.splitext(filename)[0].replace('_', ' ')))

    elif INPUT_MODE == 'CUSTOM':
        print(f"Mode: CUSTOM. Using specified files and custom titles...")
        if len(files_for_custom_mode) != len(titles_for_custom_mode):
            print("Error: The number of files and titles for CUSTOM mode must be the same.")
            sys.exit()
        for filename, title in zip(files_for_custom_mode, titles_for_custom_mode):
            chapters_to_merge.append((os.path.join(SOURCE_FOLDER, filename), title))
    else:
        print(f"Error: Invalid INPUT_MODE '{INPUT_MODE}'.")
        sys.exit()
        
    final_docx = merge_documents(chapters_to_merge, OUTPUT_DOCX_FILENAME)

    if final_docx and CREATE_PDF:
        convert_to_pdf(final_docx, OUTPUT_PDF_FILENAME)