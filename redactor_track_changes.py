#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import xml.etree.ElementTree as ET
import argparse
import re
from pathlib import Path
import zipfile
import tempfile
import shutil
from datetime import datetime


class DocumentXMLTrackChangesRedactor:
    """
    A redactor that applies Word's Track Changes markup to redacted content,
    showing deletions with strikethrough and insertions as underlined asterisks.
    
    This is a separate implementation from the regular redactor that replaces text with asterisks.
    """
    
    def __init__(self):
        # Word namespace
        self.WORD_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
        # Register namespace to maintain proper XML output
        ET.register_namespace('w', self.WORD_NS)
        
        # Track changes metadata
        self.author = "docx_tools"
        self.revision_id = 1
        self.change_date = datetime.now().strftime("%Y-%m-%dT%H:%M:%SZ")
    
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
    
    def _get_next_revision_id(self):
        """Get the next revision ID and increment counter."""
        current_id = self.revision_id
        self.revision_id += 1
        return str(current_id)
    
    def _create_track_change_runs(self, deleted_text, inserted_text):
        """
        Create deletion and insertion runs with proper track changes markup.
        
        Args:
            deleted_text: The text that was deleted
            inserted_text: The text that replaces it (asterisks)
            
        Returns:
            tuple: (deletion_run, insertion_run)
        """
        # Create deletion run
        del_run = ET.Element(f"{{{self.WORD_NS}}}r")
        
        # Add deletion markup
        del_elem = ET.SubElement(del_run, f"{{{self.WORD_NS}}}del")
        del_elem.set(f"{{{self.WORD_NS}}}id", self._get_next_revision_id())
        del_elem.set(f"{{{self.WORD_NS}}}author", self.author)
        del_elem.set(f"{{{self.WORD_NS}}}date", self.change_date)
        
        # Add deleted text using delText for proper strikethrough
        del_text = ET.SubElement(del_elem, f"{{{self.WORD_NS}}}delText")
        del_text.text = deleted_text
        
        # Create insertion run
        ins_run = ET.Element(f"{{{self.WORD_NS}}}r")
        
        # Add insertion markup
        ins_elem = ET.SubElement(ins_run, f"{{{self.WORD_NS}}}ins")
        ins_elem.set(f"{{{self.WORD_NS}}}id", self._get_next_revision_id())
        ins_elem.set(f"{{{self.WORD_NS}}}author", self.author)
        ins_elem.set(f"{{{self.WORD_NS}}}date", self.change_date)
        
        # Add inserted text
        ins_text = ET.SubElement(ins_elem, f"{{{self.WORD_NS}}}t")
        ins_text.text = inserted_text
        
        return del_run, ins_run
    
    def _apply_redactions_to_paragraph(self, paragraph, redaction_positions):
        """
        Apply all redactions to a paragraph by replacing text with track changes markup.
        
        Args:
            paragraph: The paragraph element
            redaction_positions: List of (start, end, original_text) tuples
        """
        if not redaction_positions:
            return
        
        # Get all text elements and their content
        text_elements, full_text = self._extract_text_from_paragraph(paragraph)
        
        if not text_elements or not full_text:
            return
        
        # Sort redaction positions by start position (process from end to beginning)
        redaction_positions.sort(key=lambda x: x[0], reverse=True)
        
        # Find all runs in the paragraph
        runs = paragraph.findall(f'{{{self.WORD_NS}}}r')
        
        # For each redaction, create track changes runs and insert them
        for start, end, original_text in redaction_positions:
            asterisks = '*' * len(original_text)
            del_run, ins_run = self._create_track_change_runs(original_text, asterisks)
            
            # Find the best position to insert - after the first run
            if runs:
                insert_pos = list(paragraph).index(runs[0]) + 1
                paragraph.insert(insert_pos, del_run)
                paragraph.insert(insert_pos + 1, ins_run)
            else:
                # If no runs found, append at the end
                paragraph.append(del_run)
                paragraph.append(ins_run)
        
        # Now remove the original text from text elements
        # We need to carefully handle text that may be split across elements
        for start, end, redacted_text in redaction_positions:
            self._remove_text_from_elements(text_elements, full_text, redacted_text)
    
    def _remove_text_from_elements(self, text_elements, full_text, text_to_remove):
        """
        Remove specific text from text elements, handling text that may span multiple elements.
        """
        # Find the position of the text to remove in the full text
        pos = full_text.find(text_to_remove)
        if pos == -1:
            return
        
        # Track position in the concatenated text
        current_pos = 0
        
        for text_elem in text_elements:
            if not text_elem.text:
                continue
                
            elem_text = text_elem.text
            elem_start = current_pos
            elem_end = current_pos + len(elem_text)
            
            # Check if this element contains part of the text to remove
            if (pos < elem_end and pos + len(text_to_remove) > elem_start):
                # Calculate how much of this element to remove
                remove_start = max(0, pos - elem_start)
                remove_end = min(len(elem_text), pos + len(text_to_remove) - elem_start)
                
                # Remove the text
                new_text = elem_text[:remove_start] + elem_text[remove_end:]
                text_elem.text = new_text if new_text else ""
            
            current_pos = elem_end
    
    def redact_document_xml(self, document_xml_path, strings_to_redact, output_path=None, case_sensitive=True):
        """
        Redact specified strings in a document.xml file using track changes markup.
        
        Args:
            document_xml_path (str): Path to the document.xml file
            strings_to_redact (list): List of strings to replace with track changes
            output_path (str, optional): Output path. If None, overwrites the input file
            case_sensitive (bool): Whether to perform case-sensitive matching
        
        Returns:
            bool: True if successful, False otherwise
        """
        try:
            print(f"Starting track changes redaction...")
            print(f"Input: {document_xml_path}")
            print(f"Searching for {len(strings_to_redact)} strings")
            for i, s in enumerate(strings_to_redact, 1):
                print(f"  {i}. '{s}' (length: {len(s)})")
            
            # Parse the XML file
            tree = ET.parse(document_xml_path)
            root = tree.getroot()
            
            # Track if any replacements were made
            replacements_made = False
            total_redactions = 0
            
            # Find all paragraph elements in the document
            paragraphs = root.findall('.//w:p', {'w': self.WORD_NS})
            print(f"Found {len(paragraphs)} paragraphs to process")
            
            for i, paragraph in enumerate(paragraphs):
                # Extract all text from the paragraph
                text_elements, full_text = self._extract_text_from_paragraph(paragraph)
                
                if not full_text or not text_elements:
                    continue
                
                # Find all redaction positions in this paragraph
                redaction_positions = []
                
                for string_to_redact in strings_to_redact:
                    if case_sensitive:
                        # Find all occurrences
                        start = 0
                        while True:
                            pos = full_text.find(string_to_redact, start)
                            if pos == -1:
                                break
                            redaction_positions.append((pos, pos + len(string_to_redact), string_to_redact))
                            start = pos + 1
                            total_redactions += 1
                            replacements_made = True
                            print(f"Found '{string_to_redact}' in paragraph {i} at position {pos}")
                    else:
                        # Use regex for case-insensitive replacement
                        pattern = re.compile(re.escape(string_to_redact), re.IGNORECASE)
                        for match in pattern.finditer(full_text):
                            start, end = match.span()
                            matched_text = full_text[start:end]
                            redaction_positions.append((start, end, matched_text))
                            total_redactions += 1
                            replacements_made = True
                            print(f"Found '{matched_text}' in paragraph {i} at position {start}")
                
                # Apply redactions to this paragraph
                if redaction_positions:
                    self._apply_redactions_to_paragraph(paragraph, redaction_positions)
            
            # Determine output path
            if output_path is None:
                output_path = document_xml_path
            
            # Write the modified XML back to file
            tree.write(output_path, xml_declaration=True, encoding='utf-8')
            
            if replacements_made:
                print(f"✅ Successfully applied track changes to {total_redactions} redaction(s)")
                print(f"Author: {self.author}")
                print(f"Output saved to {output_path}")
            else:
                print(f"❌ No instances of the specified strings were found")
            
            return True
            
        except Exception as e:
            print(f"❌ Error applying track changes: {str(e)}")
            import traceback
            traceback.print_exc()
            return False
    
    def redact_document_xml_case_insensitive(self, document_xml_path, strings_to_redact, output_path=None):
        """
        Redact specified strings in a document.xml file (case-insensitive matching) with track changes.
        This is now handled by the main redact_document_xml method.
        """
        return self.redact_document_xml(document_xml_path, strings_to_redact, output_path, case_sensitive=False)
    
    def redact_docx_file(self, docx_path, strings_to_redact, output_path=None, case_sensitive=True):
        """
        Redact strings in a complete .docx file using track changes markup.
        
        Args:
            docx_path (str): Path to the .docx file
            strings_to_redact (list): List of strings to redact
            output_path (str, optional): Output path for redacted .docx file
            case_sensitive (bool): Whether to perform case-sensitive matching
        
        Returns:
            bool: True if successful, False otherwise
        """
        try:
            docx_path = Path(docx_path)
            
            if output_path is None:
                output_path = docx_path.with_name(f"{docx_path.stem}_track_changes{docx_path.suffix}")
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
                
                # Redact the document.xml with track changes
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
                
                print(f"Successfully created redacted .docx file with track changes: {output_path}")
                return True
            
        except Exception as e:
            print(f"Error redacting .docx file with track changes: {str(e)}")
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
            
            print(f"Debug: Found {len(paragraphs)} paragraphs")
            
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
    """Main function to handle command line arguments and execute track changes redaction."""
    parser = argparse.ArgumentParser(description='Redact strings in Word documents using track changes markup')
    parser.add_argument('file', help='Path to document.xml or .docx file')
    parser.add_argument('strings', nargs='+', help='String(s) to redact with track changes')
    parser.add_argument('-o', '--output', help='Output file path')
    parser.add_argument('-c', '--case-insensitive', action='store_true',
                       help='Perform case-insensitive matching')
    parser.add_argument('--docx', action='store_true',
                       help='Treat input as .docx file instead of document.xml')
    parser.add_argument('--debug', action='store_true',
                       help='Show document structure (useful for debugging)')
    
    args = parser.parse_args()
    
    redactor = DocumentXMLTrackChangesRedactor()
    
    # Debug mode
    if args.debug:
        if args.docx or args.file.endswith('.docx'):
            print("Debug mode not supported for .docx files. Please extract the document.xml first.")
        else:
            redactor.debug_document_structure(args.file)
        return
    
    if args.docx or args.file.endswith('.docx'):
        # Redact entire .docx file with track changes
        success = redactor.redact_docx_file(
            args.file, 
            args.strings, 
            args.output, 
            case_sensitive=not args.case_insensitive
        )
    else:
        # Redact document.xml file only with track changes
        success = redactor.redact_document_xml(
            args.file, 
            args.strings, 
            args.output,
            case_sensitive=not args.case_insensitive
        )
    
    if success:
        print("✅ Track changes redaction completed successfully!")
    else:
        print("❌ Track changes redaction failed!")


# Example usage functions
def example_usage():
    """Examples of how to use the track changes redactor programmatically."""
    redactor = DocumentXMLTrackChangesRedactor()
    
    # Example 1: Redact strings in document.xml with track changes
    strings_to_redact = ["John Doe", "confidential", "123-45-6789"]
    redactor.redact_document_xml("output_directory/word/document.xml", strings_to_redact, 
                                "output_directory/word/document_track_changes.xml")
    
    # Example 2: Debug document structure
    redactor.debug_document_structure("output_directory/word/document.xml")
    
    # Example 3: Case-insensitive redaction with track changes
    redactor.redact_document_xml("output_directory/word/document.xml", ["CONFIDENTIAL"], 
                                case_sensitive=False)


if __name__ == "__main__":
    main()