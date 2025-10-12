# -*- coding: utf-8 -*-
"""
Merges multiple .docx files into a single document with a Table of Contents.

This script combines specified .docx files into a single Word document. 
It automatically generates a 'Contents' page that lists each chapter topic, 
alongside its starting page number.

Features:
- Flexible chapter definition: Merge all files from a folder, a specific
  list of files, or a list of files with custom titles.
- Strips any pre-existing page numbers from source files by not copying headers/footers.
- Adds new, continuous page numbering to the footer of the merged document.
- Inserts page breaks between each appended chapter.
- Optionally, converts the final .docx file to a .pdf.

Dependencies:
- python-docx: For handling .docx files.
- lxml: A dependency of python-docx.
- comtypes & pywin32 (Windows only): For PDF conversion and TOC updates.

Usage:
1. Place your .docx files in a known location (e.g., the `source_chapters` folder).
2. Go to the `if __name__ == '__main__':` block at the bottom of the script.
3. Choose your desired `INPUT_MODE` ('FOLDER', 'FILE_LIST', or 'CUSTOM').
4. Configure the corresponding lists or folder name for your chosen mode.
5. Run the script.
"""
import os
import sys
from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

# --- Helper Function for Table of Contents ---

def add_table_of_contents(doc):
    """
    Adds a table of contents field to the document.
    Word must be used to update this field to generate the actual TOC.
    """
    # Add a "Contents" heading
    doc.add_heading('Contents', level=1)
    
    paragraph = doc.add_paragraph()
    run = paragraph.add_run()
    
    # Create the TOC field
    fldChar_begin = OxmlElement('w:fldChar')
    fldChar_begin.set(qn('w:fldCharType'), 'begin')
    
    instrText = OxmlElement('w:instrText')
    instrText.set(qn('xml:space'), 'preserve')
    instrText.text = 'TOC \\o "1-3" \\h \\z \\u'
    
    fldChar_separate = OxmlElement('w:fldChar')
    fldChar_separate.set(qn('w:fldCharType'), 'separate')
    
    fldChar_end = OxmlElement('w:fldChar')
    fldChar_end.set(qn('w:fldCharType'), 'end')
    
    # Append the elements to the run
    run._r.append(fldChar_begin)
    run._r.append(instrText)
    run._r.append(fldChar_separate)
    run = paragraph.add_run('Right-click and select "Update Field" to generate the table of contents.')
    run.font.italic = True
    run = paragraph.add_run()
    run._r.append(fldChar_end)
    
    doc.add_page_break()


# --- Main Merging Function ---

def merge_documents(chapters, final_doc_path):
    """
    Merges a list of documents into a single document.
    
    Args:
        chapters (list): A list of tuples, where each tuple is (filepath, title).
        final_doc_path (str): The path to save the final merged document.
    """
    if not chapters:
        print("The list of chapters to merge is empty. Aborting.")
        return None

    master_document = Document()
    add_table_of_contents(master_document)
    
    print("Starting the merge process...")
    
    for filepath, chapter_title in chapters:
        print(f"  -> Merging '{filepath}' as '{chapter_title}'")
        
        if not os.path.exists(filepath):
            print(f"    [!] WARNING: File not found, skipping: {filepath}")
            continue

        master_document.add_heading(chapter_title, level=1)
        source_doc = Document(filepath)
        
        for element in source_doc.element.body:
            master_document.element.body.append(element)
            
        master_document.add_page_break()

    # Remove the last page break
    if master_document.paragraphs[-1].text == '':
        p = master_document.paragraphs[-1]
        p_element = p._element
        p_element.getparent().remove(p_element)

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
        if doc:
            doc.Close(SaveChanges=False)
        if word:
            word.Quit()


# --- Main Execution Block ---

if __name__ == '__main__':
    # ========================== CONFIGURATION ==========================
    
    # --- Step 1: Choose your Input Mode ---
    # Options: 'FOLDER', 'FILE_LIST', 'CUSTOM'
    INPUT_MODE = 'CUSTOM'

    # --- Step 2: Set General Parameters ---
    SOURCE_FOLDER = 'source_chapters'  # Used for all modes to find files
    OUTPUT_DOCX_FILENAME = 'merged_document.docx'
    OUTPUT_PDF_FILENAME = 'merged_document.pdf'
    
    # Set to True to enable PDF conversion (Windows with MS Word only)
    CREATE_PDF = True

    # --- Step 3: Configure the selected Input Mode ---
    # The script will use the configuration that matches the INPUT_MODE set above.
    
    # Configuration for 'FOLDER' mode
    # Merges all .docx files found in SOURCE_FOLDER in alphabetical order.
    if INPUT_MODE == 'FOLDER':
        print(f"Mode selected: FOLDER. Reading all .docx files from '{SOURCE_FOLDER}'")

    # Configuration for 'FILE_LIST' mode
    # Merges the files listed below in the specified order.
    # Chapter titles will be derived from the filenames.
    files_for_list_mode = [
        '02_System_Architecture.docx',
        '01_Introduction.docx',
        '03_Conclusion.docx',
    ]

    # Configuration for 'CUSTOM' mode
    # Merges the files listed below using the corresponding custom titles.
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
    
    # --- Script Execution ---
    
    # Create source folder if it doesn't exist
    if INPUT_MODE == 'FOLDER':
        if not os.path.exists(SOURCE_FOLDER):
            os.makedirs(SOURCE_FOLDER)
            print(f"Created a folder named '{SOURCE_FOLDER}'.")
            print("Please add your .docx chapter files into this folder and run the script again.")
            sys.exit()
    
    # Prepare the list of chapters based on the selected mode
    chapters_to_merge = []
    
    if INPUT_MODE == 'FOLDER':
        source_files = sorted([f for f in os.listdir(SOURCE_FOLDER) if f.endswith('.docx')])
        if not source_files:
             print(f"No .docx files found in '{SOURCE_FOLDER}'.")
             sys.exit()
        for filename in source_files:
            filepath = os.path.join(SOURCE_FOLDER, filename)
            title = os.path.splitext(filename)[0]
            chapters_to_merge.append((filepath, title))
            
    elif INPUT_MODE == 'FILE_LIST':
        for filename in files_for_list_mode:
            filepath = os.path.join(SOURCE_FOLDER, filename)
            title = os.path.splitext(filename)[0].replace('_', ' ')
            chapters_to_merge.append((filepath, title))

    elif INPUT_MODE == 'CUSTOM':
        if len(files_for_custom_mode) != len(titles_for_custom_mode):
            print("Error: The number of files and titles for CUSTOM mode must be the same.")
            sys.exit()
        for filename, title in zip(files_for_custom_mode, titles_for_custom_mode):
            filepath = os.path.join(SOURCE_FOLDER, filename)
            chapters_to_merge.append((filepath, title))

    else:
        print(f"Error: Invalid INPUT_MODE '{INPUT_MODE}'. Please choose 'FOLDER', 'FILE_LIST', or 'CUSTOM'.")
        sys.exit()
        
    # Merge the documents
    final_docx = merge_documents(chapters_to_merge, OUTPUT_DOCX_FILENAME)

    # Convert to PDF if requested and the docx was created
    if final_docx and CREATE_PDF:
        convert_to_pdf(final_docx, OUTPUT_PDF_FILENAME)