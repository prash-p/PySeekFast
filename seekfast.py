#!/usr/bin/env python3
"""
seekfast_clone.py - A command line tool to convert Word documents to text and search for specific words or phrases.
"""

import os
import sys
import re
import argparse
import concurrent.futures
from pathlib import Path
import docx2txt
from PyPDF2 import PdfReader
import olefile

def extract_text_from_doc(doc_path):
    """Extract text from a .doc file using olefile."""
    try:
        if not olefile.isOleFile(doc_path):
            return ""
        
        ole = olefile.OleFileIO(doc_path)
        if ole.exists('WordDocument'):
            # Extract text from the WordDocument stream
            with ole.openstream('WordDocument') as stream:
                content = stream.read()
                
            # This is a simplified extraction that works for basic .doc files
            # For more complex .doc files, libraries like antiword or textract might be needed
            text = ""
            for i in range(0, len(content), 2):
                if i+1 < len(content) and content[i+1] == 0:
                    char = chr(content[i])
                    if char.isprintable() or char.isspace():
                        text += char
            
            return text
        return ""
    except Exception as e:
        print(f"Error extracting text from .doc file {doc_path}: {e}", file=sys.stderr)
        return ""

def extract_text_from_docx(docx_path):
    """Extract text from a .docx file using docx2txt."""
    try:
        return docx2txt.process(docx_path)
    except Exception as e:
        print(f"Error extracting text from .docx file {docx_path}: {e}", file=sys.stderr)
        return ""

def extract_text_from_pdf(pdf_path):
    """Extract text from a PDF file using PyPDF2."""
    try:
        pdf = PdfReader(pdf_path)
        text = ""
        for page in pdf.pages:
            text += page.extract_text() + "\n"
        return text
    except Exception as e:
        print(f"Error extracting text from PDF file {pdf_path}: {e}", file=sys.stderr)
        return ""

def extract_text(file_path):
    """Extract text from various document formats."""
    file_path_lower = file_path.lower()
    
    if file_path_lower.endswith(".docx"):
        return extract_text_from_docx(file_path)
    elif file_path_lower.endswith(".doc"):
        return extract_text_from_doc(file_path)
    elif file_path_lower.endswith(".pdf"):
        return extract_text_from_pdf(file_path)
    else:
        # For unsupported formats, return empty string
        return ""

def search_file(file_path, search_terms, case_sensitive=False, whole_word=False, regex=False):
    """Search for terms in a file and return matching lines with context."""
    text = extract_text(file_path)
    if not text:
        return []
    
    lines = text.splitlines()
    results = []
    
    for term in search_terms:
        if regex:
            try:
                pattern = re.compile(term, 0 if case_sensitive else re.IGNORECASE)
            except re.error:
                print(f"Invalid regex pattern: {term}", file=sys.stderr)
                continue
        else:
            if whole_word:
                pattern = re.compile(r'\b' + re.escape(term) + r'\b', 0 if case_sensitive else re.IGNORECASE)
            else:
                pattern = re.compile(re.escape(term), 0 if case_sensitive else re.IGNORECASE)
        
        for i, line in enumerate(lines):
            if pattern.search(line):
                # Get context (lines before and after)
                start = max(0, i - 1)
                end = min(len(lines), i + 2)
                context = lines[start:end]
                
                results.append({
                    'file': file_path,
                    'line_number': i + 1,
                    'term': term,
                    'context': '\n'.join(context).strip()
                })
    
    return results

def find_files(directory, extensions):
    """Find all files with specific extensions in a directory and its subdirectories."""
    found_files = []
    for ext in extensions:
        # Make sure the extension starts with a dot
        if not ext.startswith('.'):
            ext = '.' + ext
        
        # Use Path.rglob to find all matching files
        found_files.extend([str(p) for p in Path(directory).rglob(f"*{ext}")])
    
    return found_files

def display_results(results, output_file=None):
    """Display search results in a readable format."""
    if not results:
        print("No matches found.")
        return
    
    output = []
    output.append(f"Found {len(results)} matches:")
    
    current_file = None
    for result in results:
        if current_file != result['file']:
            current_file = result['file']
            output.append(f"\nFile: {current_file}")
        
        output.append(f"Match for '{result['term']}' at line {result['line_number']}:")
        output.append(f"{result['context']}")
        output.append("-" * 60)
    
    output_text = "\n".join(output)
    
    if output_file:
        with open(output_file, 'w', encoding='utf-8') as f:
            f.write(output_text)
        print(f"Results saved to {output_file}")
    else:
        print(output_text)

def main():
    parser = argparse.ArgumentParser(description='Search for terms in Word documents and other file types.')
    parser.add_argument('directory', help='Directory to search in')
    parser.add_argument('terms', nargs='+', help='Terms to search for')
    parser.add_argument('-e', '--extensions', default=['doc', 'docx', 'pdf'], nargs='+',
                        help='File extensions to search (default: doc, docx, pdf)')
    parser.add_argument('-c', '--case-sensitive', action='store_true', help='Case sensitive search')
    parser.add_argument('-w', '--whole-word', action='store_true', help='Match whole words only')
    parser.add_argument('-r', '--regex', action='store_true', help='Interpret search terms as regular expressions')
    parser.add_argument('-o', '--output', help='Save results to output file')
    parser.add_argument('-j', '--jobs', type=int, default=os.cpu_count(),
                        help='Number of parallel jobs (default: number of CPU cores)')
    args = parser.parse_args()
    
    # Find all files with the specified extensions
    files = find_files(args.directory, args.extensions)
    
    if not files:
        print(f"No files with extensions {args.extensions} found in {args.directory}")
        return
    
    print(f"Searching {len(files)} files for {len(args.terms)} terms...")
    
    # Process files in parallel
    all_results = []
    with concurrent.futures.ProcessPoolExecutor(max_workers=args.jobs) as executor:
        future_to_file = {
            executor.submit(search_file, file, args.terms, args.case_sensitive, args.whole_word, args.regex): file
            for file in files
        }
        for future in concurrent.futures.as_completed(future_to_file):
            file = future_to_file[future]
            try:
                results = future.result()
                all_results.extend(results)
            except Exception as e:
                print(f"Error processing {file}: {e}", file=sys.stderr)
    
    # Sort results by filename and line number
    all_results.sort(key=lambda x: (x['file'], x['line_number']))
    
    # Display the results
    display_results(all_results, args.output)

if __name__ == "__main__":
    main()
