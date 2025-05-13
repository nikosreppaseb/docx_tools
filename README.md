# DOCX Tools

This directory contains Python scripts for working with `.docx` files.

## Overview

The tools provided are:
- **`converter.py`**: Converts `.docx` files to and from the OpenXML format (a directory structure of XML files). This is useful for inspecting or programmatically modifying the underlying XML content of a Word document.
- **`redactor.py`**: Redacts specified strings within a `.docx` file or its `document.xml` component. It replaces target strings with asterisks (`*`) while preserving XML structure and formatting, even when text is split across multiple XML elements.

## `converter.py`

### Purpose

The `DOCXOpenXMLConverter` class in [`converter.py`](docx_tools/converter.py:9) allows for lossless conversion between the `.docx` format and its underlying OpenXML structure. Since `.docx` files are essentially ZIP archives containing XML files and other resources, this script facilitates extracting these components for easier manipulation and then re-packaging them.

### Features

- **DOCX to OpenXML**: Extracts the contents of a `.docx` file into a specified directory, preserving the OpenXML structure.
- **OpenXML to DOCX**: Compresses an OpenXML directory structure back into a `.docx` file.
- **XML Prettification**: Automatically pretty-prints XML files (`.xml`, `.rels`) upon extraction for improved readability.
- **Structure Validation**: Includes a method to validate the basic structure of a `.docx` file (e.g., presence of `[Content_Types].xml`, `_rels/.rels`, `word/document.xml`).

### Usage

#### Command-Line Interface

The script can be run from the command line:

```bash
python docx_tools/converter.py <operation> <input> <output> [--validate]
```

- **`operation`**:
    - `to-openxml`: Convert a `.docx` file to an OpenXML directory.
    - `to-docx`: Convert an OpenXML directory to a `.docx` file.
- **`input`**:
    - Path to the input `.docx` file (for `to-openxml`).
    - Path to the input OpenXML directory (for `to-docx`).
- **`output`**:
    - Path to the output directory for OpenXML files (for `to-openxml`).
    - Path for the output `.docx` file (for `to-docx`).
- **`--validate`** (optional, for `to-openxml`): Validate the `.docx` file structure before conversion.

**Examples:**

1.  **Convert DOCX to OpenXML:**
    ```bash
    python docx_tools/converter.py to-openxml mydocument.docx mydocument_openxml
    ```

2.  **Convert OpenXML to DOCX:**
    ```bash
    python docx_tools/converter.py to-docx mydocument_openxml mydocument_converted.docx
    ```

3.  **Validate and Convert DOCX to OpenXML:**
    ```bash
    python docx_tools/converter.py to-openxml mydocument.docx mydocument_openxml --validate
    ```

#### Programmatic Usage

```python
from docx_tools.converter import DOCXOpenXMLConverter

converter = DOCXOpenXMLConverter()

# Convert DOCX to OpenXML
docx_file = "example.docx"  # Replace with your .docx file
openxml_dir = "example_openxml"
if converter.docx_to_openxml(docx_file, openxml_dir):
    print(f"Successfully converted {docx_file} to {openxml_dir}")

    # Convert OpenXML back to DOCX
    new_docx_file = "example_reverted.docx"
    if converter.openxml_to_docx(openxml_dir, new_docx_file):
        print(f"Successfully converted {openxml_dir} back to {new_docx_file}")

# Validate a DOCX file
if converter.validate_docx_structure("another_document.docx"):
    print("another_document.docx has a valid structure.")
else:
    print("another_document.docx structure is invalid or file is corrupted.")
```

## `redactor.py`

### Purpose

The `DocumentXMLRedactor` class in [`redactor.py`](docx_tools/redactor.py:10) is designed to find and replace specified text strings within the `word/document.xml` part of a `.docx` file, or directly within a `.docx` file itself. The primary use case is redacting sensitive information by replacing it with asterisks (`*`). It's built to handle cases where the text to be redacted might be split across multiple `<w:t>` (text run) XML elements, ensuring that formatting is preserved.

### Features

- **Targeted Redaction**: Replaces specified strings with a sequence of asterisks of the same length.
- **Handles Split Text**: Correctly redacts text even if it's fragmented across multiple `<w:t>` elements within a paragraph.
- **Preserves Formatting**: Modifies only the text content, leaving XML structure and associated formatting intact.
- **Case Sensitivity**: Supports both case-sensitive and case-insensitive redaction.
- **Operates on `document.xml` or `.docx`**: Can redact a standalone `document.xml` file or an entire `.docx` package.
- **Temporary File Handling**: Uses a temporary directory when processing `.docx` files to avoid altering the original until the operation is complete.
- **Debugging**: Includes a method to print the structure of paragraphs and their text elements, aiding in understanding how text is stored.

### Usage

#### Command-Line Interface

The script can be run from the command line:

```bash
python docx_tools/redactor.py <file> <string1> [<string2> ...] [-o <output>] [-c] [--docx] [--debug]
```

- **`file`**: Path to the `document.xml` file or `.docx` file.
- **`strings`**: One or more strings to redact.
- **`-o, --output <output>`** (optional): Path for the output redacted file. If not provided, the input `.docx` file will be overwritten with a `_redacted` suffix, or the `document.xml` file will be overwritten directly.
- **`-c, --case-insensitive`** (optional): Perform case-insensitive matching. Default is case-sensitive.
- **`--docx`** (optional): Explicitly treat the input file as a `.docx` file. The script also infers this if the file ends with `.docx`.
- **`--debug`** (optional): Show the document structure (first 5 paragraphs of `document.xml`) for debugging purposes. Not for `.docx` files directly.

**Examples:**

1.  **Redact strings in `document.xml` (case-sensitive):**
    ```bash
    python docx_tools/redactor.py path/to/your/document.xml "Secret Phrase" " Confidential Info "
    ```

2.  **Redact strings in a `.docx` file (case-insensitive) and save to a new file:**
    ```bash
    python docx_tools/redactor.py mydocument.docx "Client Name" "Project Alpha" -o mydocument_redacted.docx -c
    ```
    Alternatively, using `--docx`:
    ```bash
    python docx_tools/redactor.py mydocument.docx "Client Name" "Project Alpha" -o mydocument_redacted.docx -c --docx
    ```

3.  **Debug `document.xml` structure:**
    ```bash
    python docx_tools/redactor.py path/to/your/document.xml --debug
    ```

#### Programmatic Usage

```python
from docx_tools.redactor import DocumentXMLRedactor

redactor = DocumentXMLRedactor()
strings_to_redact = ["John Doe", "Project Phoenix", "SSN: 123-45-6789"]

# Example 1: Redact a document.xml file
# Assume 'source_openxml/word/document.xml' is the path to your document.xml
# and 'redacted_openxml/word/document.xml' is where you want to save it.
# Ensure the output directory exists or handle its creation.
# For simplicity, this example assumes it's run after converter.py has extracted a docx.
# Path("redacted_openxml/word").mkdir(parents=True, exist_ok=True) # Example directory creation
# redactor.redact_document_xml(
# "source_openxml/word/document.xml",
# strings_to_redact,
# "redacted_openxml/word/document_redacted.xml",
# case_sensitive=True
# )

# Example 2: Redact an entire .docx file
input_docx = "mydocument.docx" # Replace with your .docx file
output_docx = "mydocument_redacted.docx"
success = redactor.redact_docx_file(
    input_docx,
    strings_to_redact,
    output_docx,
    case_sensitive=False # Case-insensitive redaction
)
if success:
    print(f"Successfully redacted {input_docx} and saved to {output_docx}")

# Example 3: Debug document structure (using an extracted document.xml)
# redactor.debug_document_structure("source_openxml/word/document.xml")
```

## How to Execute

For detailed instructions on setting up the environment and running these scripts, please refer to [`how_to_execute.md`](docx_tools/how_to_execute.md:1).