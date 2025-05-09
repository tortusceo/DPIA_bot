import os
from dpia_bot import convert_markdown_to_original_format

# Sample Markdown content for testing
sample_markdown = """
# Test Document

This is a **test** paragraph with some *italic* text.

## Subsection

- Item 1
- Item 2
  - Sub-item 2.1

Another paragraph.

| Header 1 | Header 2 |
|----------|----------|
| Cell 1   | Cell 2   |
| Cell 3   | Cell 4   |
"""

output_docx_filename = "test_pandoc_output.docx"

print(f"Attempting to convert sample Markdown to {output_docx_filename} using Pandoc...")

# Ensure the function can be called. If it relies on other parts of dpia_bot for setup 
# (though this one should be fairly self-contained for .docx), this might need adjustment.
result_file = convert_markdown_to_original_format(sample_markdown, ".docx", output_docx_filename)

if os.path.exists(result_file) and result_file == output_docx_filename:
    print(f"Test successful! Output saved to: {result_file}")
    print("Please open and check the formatting of the .docx file.")
else:
    print(f"Test failed or an error occurred. Result: {result_file}")
    if not os.path.exists(output_docx_filename):
        print(f"{output_docx_filename} was not created.")
    elif result_file != output_docx_filename:
        print(f"Output file was {result_file} instead of expected {output_docx_filename}") 