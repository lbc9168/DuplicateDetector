# -*- coding: utf-8 -*-
"""
Created on Fri Feb 21 10:38:48 2025

@author: Bingcai Liu
"""

from docx import Document



def extract_text_from_docx(file_path):
    """Extract text from a Word document (.docx)"""
    doc = Document(file_path)
    return "\n".join([para.text for para in doc.paragraphs])

# Replace the docx_path with the path to your document
file_path_1 = "C:/Users/User/OneDrive/IIGF/Main Projects/202410_TNC DNS_CIFF DDS/Writing/TNC-DNS/plag-checking-python project/2024-TNC-ver1.docx"
file_path_2 = "C:/Users/User/OneDrive/IIGF/Main Projects/202410_TNC DNS_CIFF DDS/Writing/TNC-DNS/plag-checking-python project/2025-TNC-ver2.docx"

try:
    text_1 = extract_text_from_docx(file_path_1)
    text_2 = extract_text_from_docx(file_path_2)

    print("Length of text_1:", len(text_1))
    print("Length of text_2:", len(text_2))
        
    if len(text_1) == 0 or len(text_2) == 0:
        print("Warning: One or both documents appear to be empty!")
        
    print("\nFirst document text preview:\n", text_1[:500])
    print("\nSecond document text preview:\n", text_2[:500])
except Exception as e:
    print(f"Error reading files: {str(e)}")
    exit(1)

# Generate an HTML comparison report

import difflib


# Compute the differences
differ = difflib.Differ()
diff = list(differ.compare(text_1.splitlines(), text_2.splitlines()))

# Filter out unchanged lines to focus on similarities and slight variations
similar_lines = [line for line in diff if line.startswith("  ") or line.startswith("? ")]
modified_lines = [line for line in diff if line.startswith("- ") or line.startswith("+ ")]

# Compute the line-by-line comparison

html_header_wrapped = """
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Debt Nature Swap Reports Comparison</title>
    <style>
        body { font-family: Arial, sans-serif; margin: 20px; padding: 20px; background-color: #f9f9f9; }
        h2 { text-align: center; color: #333; }
        table { 
            width: 100%; 
            border-collapse: collapse; 
            margin-top: 20px; 
            table-layout: fixed; 
        }
        
        /* Target specific columns by their position */
        td:nth-child(1) { width: 3%; }    /* Navigation column 1 */
        td:nth-child(2) { width: 3%; }    /* First line number column */
        td:nth-child(3) { width: 44%; }   /* First content column */
        td:nth-child(4) { width: 3%; }    /* Navigation column 2 */
        td:nth-child(5) { width: 3%; }    /* Second line number column */
        td:nth-child(6) { width: 44%; }   /* Second content column */
        
        th, td { 
            padding: 10px; 
            border: 1px solid #ccc; 
            text-align: left; 
            word-wrap: break-word; 
            overflow-wrap: break-word; 
            white-space: pre-wrap; 
        }
        th { background-color: #0073e6; color: white; }
        .diff_add { background-color: #d4fcbc; }
        .diff_sub { background-color: #fbb4b4; }
        .diff_change { background-color: #fff2cc; }
        pre { 
            white-space: pre-wrap; 
            font-family: monospace; 
            word-wrap: break-word; 
            overflow-wrap: break-word; 
        }
    </style>
</head>
<body>
    <h2>Debt for Nature Swap Reports: Side-by-Side Comparison</h2>
"""
html_footer = """
</body>
</html>
"""

# Generate enhanced HTML diff
differ = difflib.HtmlDiff()
enhanced_html_diff = differ.make_table(text_1.splitlines(), text_2.splitlines(), 
                                       fromdesc="Report 1 (2024-IIGF)", todesc="Report 2 (2025-TNC)", 
                                       context=True, numlines=5)

# Combine header, diff, and footer
final_html_output = html_header_wrapped + enhanced_html_diff + html_footer

# Save the final HTML output to a file
output_file = "comparison_report.html"
try:
    with open(output_file, "w", encoding='utf-8') as f:
        f.write(final_html_output)
    print(f"\nComparison report successfully saved to {output_file}")
    print(f"HTML file size: {len(final_html_output)} characters")
except Exception as e:
    print(f"Error saving HTML file: {str(e)}")

