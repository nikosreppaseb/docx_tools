Converter Usage

Command Line Interface

# Convert DOCX to OpenXML
python converter.py to-openxml input.docx output_directory
# Convert OpenXML back to DOCX
python converter.py to-docx openxml_directory output.docx
# With validation
python converter.py to-openxml input.docx output_directory --validate

Programmatic Usage

converter = DOCXOpenXMLConverter()
# DOCX to OpenXML
converter.docx_to_openxml("document.docx", "openxml_folder")
# OpenXML to DOCX
converter.openxml_to_docx("openxml_folder", "converted.docx")
===========================================================================================================

Redactor Usage

Command Line Interface

# Redact specific strings in document.xml
python redactor.py document.xml "John Doe" "confidential" "secret"
# Case-insensitive redaction
python redactor.py document.xml "CONFIDENTIAL" -c
# Redact in a complete .docx file
python redactor.py document.docx "sensitive_data" --docx -o redacted_document.docx
# Specify output file
python redactor.py document.xml "password" -o document_redacted.xml

Programmatic Usage

redactor = DocumentXMLRedactor()
# Redact strings in document.xml
strings_to_redact = ["John Doe", "confidential", "secret"]
redactor.redact_document_xml("word/document.xml", strings_to_redact)

# Redact a complete .docx file with case-insensitive matching
redactor.redact_docx_file("document.docx", strings_to_redact, 
                         "redacted.docx", case_sensitive=False)