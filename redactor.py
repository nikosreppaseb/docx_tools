import xml.etree.ElementTree as ET
import argparse
import re
from pathlib import Path
import zipfile
import tempfile
import shutil


class DocumentXMLRedactor:
    """
    A class to redact (replace with stars) specified strings in a Word document.xml file
    while preserving the XML structure and formatting, handling text split across multiple <w:t> elements.
    """
    
    def __init__(self):
        # Word namespace
        self.WORD_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
        # Register namespace to maintain proper XML output
        ET.register_namespace('w', self.WORD_NS)
    
    def _extract_text_from_paragraph(self, paragraph_elem):
        """
        Extract all text from a paragraph element, maintaining the order of <w:t> elements.
        
        Args:
            paragraph_elem: The paragraph XML element
            
        Returns:
            tuple: (list_of_text_elements, concatenated_text)
        """
        text_elements = []
        text_parts = []
        
        # Find all text elements within the paragraph
        for elem in paragraph_elem.iter():
            if elem.tag == f"{{{self.WORD_NS}}}t":
                text_elements.append(elem)
                text_parts.append(elem.text or "")
        
        return text_elements, ''.join(text_parts)
    
    def _redact_text_in_elements(self, text_elements, original_text, redacted_text):
        """
        Apply redaction to the original text elements while preserving formatting.
        
        Args:
            text_elements: List of <w:t> elements
            original_text: The concatenated original text
            redacted_text: The text with redactions applied
        """
        if not text_elements or original_text == redacted_text:
            return
        
        # Track position in the redacted text
        redacted_pos = 0
        original_pos = 0
        
        for text_elem in text_elements:
            elem_text = text_elem.text or ""
            elem_length = len(elem_text)
            
            if elem_length == 0:
                continue
            
            # Extract the corresponding portion from the redacted text
            redacted_portion = redacted_text[redacted_pos:redacted_pos + elem_length]
            
            # Update the element's text
            text_elem.text = redacted_portion
            
            # Move positions forward
            redacted_pos += elem_length
            original_pos += elem_length
    
    def redact_document_xml(self, document_xml_path, strings_to_redact, output_path=None, case_sensitive=True):
        """
        Redact specified strings in a document.xml file by replacing them with stars.
        Handles text split across multiple <w:t> elements.
        
        Args:
            document_xml_path (str): Path to the document.xml file
            strings_to_redact (list): List of strings to replace with stars
            output_path (str, optional): Output path. If None, overwrites the input file
            case_sensitive (bool): Whether to perform case-sensitive matching
        
        Returns:
            bool: True if successful, False otherwise
        """
        try:
            # Parse the XML file
            tree = ET.parse(document_xml_path)
            root = tree.getroot()
            
            # Track if any replacements were made
            replacements_made = False
            
            # Find all paragraph elements in the document
            paragraphs = root.findall('.//w:p', {'w': self.WORD_NS})
            
            for paragraph in paragraphs:
                # Extract all text from the paragraph
                text_elements, full_text = self._extract_text_from_paragraph(paragraph)
                
                if not full_text or not text_elements:
                    continue
                
                # Apply redactions to the full text
                redacted_text = full_text
                
                for string_to_redact in strings_to_redact:
                    if case_sensitive:
                        # Simple string replacement for case-sensitive
                        if string_to_redact in redacted_text:
                            replacement = '*' * len(string_to_redact)
                            redacted_text = redacted_text.replace(string_to_redact, replacement)
                            replacements_made = True
                    else:
                        # Use regex for case-insensitive replacement
                        pattern = re.compile(re.escape(string_to_redact), re.IGNORECASE)
                        matches = list(pattern.finditer(redacted_text))
                        
                        # Replace matches from end to beginning to maintain positions
                        for match in reversed(matches):
                            start, end = match.span()
                            matched_text = redacted_text[start:end]
                            replacement = '*' * len(matched_text)
                            redacted_text = redacted_text[:start] + replacement + redacted_text[end:]
                            replacements_made = True
                
                # Apply the redactions back to the individual elements
                self._redact_text_in_elements(text_elements, full_text, redacted_text)
            
            # Determine output path
            if output_path is None:
                output_path = document_xml_path
            
            # Write the modified XML back to file
            tree.write(output_path, xml_declaration=True, encoding='utf-8')
            
            if replacements_made:
                print(f"Successfully redacted strings in {document_xml_path}")
                print(f"Output saved to {output_path}")
            else:
                print(f"No instances of the specified strings were found in {document_xml_path}")
            
            return True
            
        except Exception as e:
            print(f"Error redacting document.xml: {str(e)}")
            import traceback
            traceback.print_exc()
            return False
    
    def redact_document_xml_case_insensitive(self, document_xml_path, strings_to_redact, output_path=None):
        """
        Redact specified strings in a document.xml file (case-insensitive matching).
        This is now handled by the main redact_document_xml method.
        """
        return self.redact_document_xml(document_xml_path, strings_to_redact, output_path, case_sensitive=False)
    
    def redact_docx_file(self, docx_path, strings_to_redact, output_path=None, case_sensitive=True):
        """
        Redact strings in a complete .docx file by extracting, modifying, and repackaging.
        
        Args:
            docx_path (str): Path to the .docx file
            strings_to_redact (list): List of strings to replace with stars
            output_path (str, optional): Output path for redacted .docx file
            case_sensitive (bool): Whether to perform case-sensitive matching
        
        Returns:
            bool: True if successful, False otherwise
        """
        try:
            docx_path = Path(docx_path)
            
            if output_path is None:
                output_path = docx_path.with_name(f"{docx_path.stem}_redacted{docx_path.suffix}")
            else:
                output_path = Path(output_path)
            
            # Create temporary directory for extraction
            with tempfile.TemporaryDirectory() as temp_dir:
                temp_dir = Path(temp_dir)
                
                # Extract the .docx file
                with zipfile.ZipFile(docx_path, 'r') as zip_ref:
                    zip_ref.extractall(temp_dir)
                
                # Locate document.xml
                document_xml_path = temp_dir / "word" / "document.xml"
                
                if not document_xml_path.exists():
                    print(f"Error: document.xml not found in {docx_path}")
                    return False
                
                # Redact the document.xml
                success = self.redact_document_xml(document_xml_path, strings_to_redact, 
                                                 case_sensitive=case_sensitive)
                
                if not success:
                    return False
                
                # Repackage the .docx file
                with zipfile.ZipFile(output_path, 'w', zipfile.ZIP_DEFLATED) as zip_ref:
                    for file_path in temp_dir.rglob('*'):
                        if file_path.is_file():
                            arc_name = file_path.relative_to(temp_dir)
                            zip_ref.write(file_path, arc_name)
                
                print(f"Successfully created redacted .docx file: {output_path}")
                return True
            
        except Exception as e:
            print(f"Error redacting .docx file: {str(e)}")
            import traceback
            traceback.print_exc()
            return False
    
    def debug_document_structure(self, document_xml_path):
        """
        Debug method to show how text is structured in the document.
        Useful for understanding text splitting across elements.
        
        Args:
            document_xml_path (str): Path to the document.xml file
        """
        try:
            tree = ET.parse(document_xml_path)
            root = tree.getroot()
            
            paragraphs = root.findall('.//w:p', {'w': self.WORD_NS})
            
            for i, paragraph in enumerate(paragraphs[:5]):  # Show first 5 paragraphs
                print(f"\n--- Paragraph {i+1} ---")
                text_elements, full_text = self._extract_text_from_paragraph(paragraph)
                print(f"Full text: '{full_text}'")
                print(f"Text elements ({len(text_elements)}):")
                for j, elem in enumerate(text_elements):
                    print(f"  {j+1}: '{elem.text or ''}'")
                    
        except Exception as e:
            print(f"Error debugging document structure: {str(e)}")


def main():
    """Main function to handle command line arguments and execute redaction."""
    parser = argparse.ArgumentParser(description='Redact strings in Word document.xml or .docx files')
    parser.add_argument('file', help='Path to document.xml or .docx file')
    parser.add_argument('strings', nargs='+', help='String(s) to redact')
    parser.add_argument('-o', '--output', help='Output file path')
    parser.add_argument('-c', '--case-insensitive', action='store_true',
                       help='Perform case-insensitive matching')
    parser.add_argument('--docx', action='store_true',
                       help='Treat input as .docx file instead of document.xml')
    parser.add_argument('--debug', action='store_true',
                       help='Show document structure (useful for debugging)')
    
    args = parser.parse_args()
    
    redactor = DocumentXMLRedactor()
    
    # Debug mode
    if args.debug:
        if args.docx or args.file.endswith('.docx'):
            print("Debug mode not supported for .docx files. Please extract the document.xml first.")
        else:
            redactor.debug_document_structure(args.file)
        return
    
    if args.docx or args.file.endswith('.docx'):
        # Redact entire .docx file
        success = redactor.redact_docx_file(
            args.file, 
            args.strings, 
            args.output, 
            case_sensitive=not args.case_insensitive
        )
    else:
        # Redact document.xml file only
        success = redactor.redact_document_xml(
            args.file, 
            args.strings, 
            args.output,
            case_sensitive=not args.case_insensitive
        )
    
    if success:
        print("Redaction completed successfully!")
    else:
        print("Redaction failed!")


# Example usage functions
def example_usage():
    """Examples of how to use the redactor programmatically."""
    redactor = DocumentXMLRedactor()
    
    # Example 1: Redact strings in document.xml (handles split text)
    strings_to_redact = ["John Doe", "confidential", "123-45-6789"]
    redactor.redact_document_xml("output_directory/word/document.xml", strings_to_redact, 
                                "output_directory/word/document_redacted.xml")
    
    # Example 2: Debug document structure
    redactor.debug_document_structure("output_directory/word/document.xml")
    
    # Example 3: Case-insensitive redaction
    redactor.redact_document_xml("output_directory/word/document.xml", ["CONFIDENTIAL"], 
                                case_sensitive=False)


if __name__ == "__main__":
    main()