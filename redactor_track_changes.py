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
    
    def _normalize_text(self, text):
        """Normalize text by removing extra whitespace and normalizing Unicode."""
        if not text:
            return ""
        # Normalize Unicode and strip/collapse whitespace
        return re.sub(r'\s+', ' ', text.strip())
    
    def _extract_text_runs_from_paragraph(self, paragraph_elem):
        """
        Extract all runs and their text from a paragraph element, maintaining order and structure.
        
        Args:
            paragraph_elem: The paragraph XML element
            
        Returns:
            list: List of (run_element, text_elements, run_text) tuples
        """
        runs_info = []
        
        # Find all run elements within the paragraph
        for run in paragraph_elem.findall(f'.//w:r', {'w': self.WORD_NS}):
            text_elements = run.findall(f'.//w:t', {'w': self.WORD_NS})
            run_text = ''.join(elem.text or '' for elem in text_elements)
            runs_info.append((run, text_elements, run_text))
        
        return runs_info
    
    def _extract_text_from_paragraph(self, paragraph_elem):
        """
        Extract all text from a paragraph element, maintaining the order of text elements.
        
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
    
    def _create_track_change_runs(self, deleted_text, inserted_text, preserve_formatting=None):
        """
        Create deletion and insertion runs with proper track changes markup.
        
        Args:
            deleted_text: The text that was deleted
            inserted_text: The text that replaces it (asterisks)
            preserve_formatting: Original run element to copy formatting from
            
        Returns:
            tuple: (deletion_run, insertion_run)
        """
        # Create deletion run
        del_run = ET.Element(f"{{{self.WORD_NS}}}r")
        
        # Copy formatting from original run if available
        if preserve_formatting is not None:
            rPr = preserve_formatting.find(f'{{{self.WORD_NS}}}rPr')
            if rPr is not None:
                del_run.append(rPr)
        
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
        
        # Copy formatting from original run if available
        if preserve_formatting is not None:
            rPr = preserve_formatting.find(f'{{{self.WORD_NS}}}rPr')
            if rPr is not None:
                ins_run.append(rPr)
        
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
        Apply all redactions to a paragraph by replacing text with track changes markup in place.
        
        Args:
            paragraph: The paragraph element
            redaction_positions: List of (start, end, original_text) tuples
        """
        if not redaction_positions:
            return
        
        # Get all text elements and their content
        text_elements, full_text = self._extract_text_from_paragraph(paragraph)
        runs_info = self._extract_text_runs_from_paragraph(paragraph)
        
        if not text_elements or not full_text or not runs_info:
            return
        
        # Sort redaction positions by start position (process from end to beginning)
        redaction_positions.sort(key=lambda x: x[0], reverse=True)
        
        # Track position in the full text
        text_position = 0
        run_positions = []
        
        # Map each run to its position in the full text
        for run, text_elems, run_text in runs_info:
            start_pos = text_position
            end_pos = text_position + len(run_text)
            run_positions.append((run, text_elems, run_text, start_pos, end_pos))
            text_position = end_pos
        
        # Process each redaction
        for start, end, original_text in redaction_positions:
            asterisks = '*' * len(original_text)
            
            # Find which run(s) contain this text
            affected_runs = []
            for run, text_elems, run_text, run_start, run_end in run_positions:
                if start < run_end and end > run_start:
                    affected_runs.append((run, text_elems, run_text, run_start, run_end))
            
            if not affected_runs:
                continue
            
            # Handle the case where text spans multiple runs
            if len(affected_runs) == 1:
                # Simple case: text is within a single run
                run, text_elems, run_text, run_start, run_end = affected_runs[0]
                
                # Calculate relative positions within the run
                rel_start = start - run_start
                rel_end = end - run_start
                
                # Split the run text
                before_text = run_text[:rel_start]
                redacted_text = run_text[rel_start:rel_end]
                after_text = run_text[rel_end:]
                
                # Create track changes runs
                del_run, ins_run = self._create_track_change_runs(redacted_text, asterisks, run)
                
                # Get the position of this run in the paragraph
                run_index = list(paragraph).index(run)
                
                # Remove the original run
                paragraph.remove(run)
                
                # Insert the replacement runs and text
                if before_text:
                    # Create run for text before redaction
                    before_run = ET.Element(f"{{{self.WORD_NS}}}r")
                    # Copy formatting
                    rPr = run.find(f'{{{self.WORD_NS}}}rPr')
                    if rPr is not None:
                        before_run.append(rPr)
                    before_t = ET.SubElement(before_run, f"{{{self.WORD_NS}}}t")
                    before_t.text = before_text
                    paragraph.insert(run_index, before_run)
                    run_index += 1
                
                # Insert track changes runs
                paragraph.insert(run_index, del_run)
                paragraph.insert(run_index + 1, ins_run)
                run_index += 2
                
                if after_text:
                    # Create run for text after redaction
                    after_run = ET.Element(f"{{{self.WORD_NS}}}r")
                    # Copy formatting
                    rPr = run.find(f'{{{self.WORD_NS}}}rPr')
                    if rPr is not None:
                        after_run.append(rPr)
                    after_t = ET.SubElement(after_run, f"{{{self.WORD_NS}}}t")
                    after_t.text = after_text
                    paragraph.insert(run_index, after_run)
                
                # Update run positions for subsequent redactions
                # (This is approximate since we're processing in reverse order)
                run_positions = [(r, te, rt, rs, re) for r, te, rt, rs, re in run_positions if r != run]
                
            else:
                # Complex case: text spans multiple runs
                # For now, let's handle this by creating the track changes after the first affected run
                first_run = affected_runs[0][0]
                run_index = list(paragraph).index(first_run)
                
                # Create track changes runs
                del_run, ins_run = self._create_track_change_runs(original_text, asterisks, first_run)
                
                # Insert after the first run
                paragraph.insert(run_index + 1, del_run)
                paragraph.insert(run_index + 2, ins_run)
                
                # Remove the text from all affected runs
                for run, text_elems, run_text, run_start, run_end in affected_runs:
                    # Calculate what part of this run to remove
                    remove_start = max(0, start - run_start)
                    remove_end = min(len(run_text), end - run_start)
                    
                    if remove_start < remove_end:
                        # Remove text from this run's text elements
                        text_to_remove = run_text[remove_start:remove_end]
                        self._remove_text_from_run(run, text_to_remove, remove_start, remove_end)
    
    def _remove_text_from_run(self, run, text_to_remove, start_pos, end_pos):
        """
        Remove specific text from a run's text elements.
        """
        text_elements = run.findall(f'.//w:t', {'w': self.WORD_NS})
        
        if not text_elements:
            return
        
        # Concatenate all text in the run
        run_text = ''.join(elem.text or '' for elem in text_elements)
        
        # Build new text
        new_text = run_text[:start_pos] + run_text[end_pos:]
        
        # Clear all text elements and put new text in first one
        for i, elem in enumerate(text_elements):
            if i == 0:
                elem.text = new_text
            else:
                elem.text = ""
    
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
                
                # Print paragraph text for debugging (first few paragraphs)
                if i < 5:
                    print(f"Paragraph {i}: '{full_text[:100]}...'")
                
                # Find all redaction positions in this paragraph
                redaction_positions = []
                
                for string_to_redact in strings_to_redact:
                    # Normalize both the search string and the text
                    normalized_search = self._normalize_text(string_to_redact)
                    normalized_text = self._normalize_text(full_text)
                    
                    if case_sensitive:
                        # Use both original and normalized text for searching
                        # First try exact match
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
                        
                        # If no exact match, try with normalized text
                        if not redaction_positions and normalized_search in normalized_text:
                            # Find the position in the normalized text
                            norm_pos = normalized_text.find(normalized_search)
                            # Convert back to original text position (approximate)
                            # This is tricky due to whitespace normalization
                            actual_pos = full_text.find(string_to_redact.strip())
                            if actual_pos != -1:
                                redaction_positions.append((actual_pos, actual_pos + len(string_to_redact.strip()), string_to_redact.strip()))
                                total_redactions += 1
                                replacements_made = True
                                print(f"Found normalized '{string_to_redact}' in paragraph {i} at position {actual_pos}")
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
                print("Debug: Let's check what text exists in the document...")
                # Debug output
                for i, paragraph in enumerate(paragraphs[:3]):
                    text_elements, full_text = self._extract_text_from_paragraph(paragraph)
                    if full_text.strip():
                        print(f"  Paragraph {i}: '{full_text[:200]}...'")
            
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