# PySeekFast
Find terms in all word documents and pdfs in a given directory!

### Installation

1. Make sure you have Python 3.6+ installed
2. Install the required dependencies:
   ```
   pip install -r requirements.txt
   ```

### Features

- Converts .doc, .docx, and PDF files to text
- Searches for multiple terms simultaneously
- Supports case-sensitive search
- Whole word matching option
- Regular expression search capability
- Multi-threaded processing for faster searches
- Context display (shows lines before and after matches)
- Output results to file

### Usage Examples

1. Basic search for a term in a directory:
   ```
   python seekfast_clone.py /path/to/documents "search term"
   ```

2. Search for multiple terms:
   ```
   python seekfast_clone.py /path/to/documents "term1" "term2" "term3"
   ```

3. Case-sensitive search for whole words:
   ```
   python seekfast_clone.py /path/to/documents "Important" -c -w
   ```

4. Search using regular expressions:
   ```
   python seekfast_clone.py /path/to/documents "import.*data" -r
   ```

5. Save results to a file:
   ```
   python seekfast_clone.py /path/to/documents "search term" -o results.txt
   ```

6. Specify file extensions to search:
   ```
   python seekfast_clone.py /path/to/documents "search term" -e doc docx pdf txt
   ```

The program will display the matching lines with some context (lines before and after), making it easier to understand the context of the matches.
