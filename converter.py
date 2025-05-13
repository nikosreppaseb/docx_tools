import zipfile
import os
import shutil
import argparse
from pathlib import Path
import xml.etree.ElementTree as ET
from xml.dom import minidom

class DOCXOpenXMLConverter:
    """
    A class to convert between .docx and OpenXML formats without losing content or formatting.
    
    .docx files are essentially renamed .zip files containing OpenXML files.
    This converter extracts/compresses the OpenXML structure while preserving all formatting.
    """
    
    def __init__(self):
        self.temp_dir = None
    
    def docx_to_openxml(self, docx_path, output_dir):
        """
        Convert a .docx file to OpenXML directory structure.
        
        Args:
            docx_path (str): Path to the input .docx file
            output_dir (str): Directory where OpenXML files will be extracted
        
        Returns:
            bool: True if successful, False otherwise
        """
        try:
            docx_path = Path(docx_path)
            output_dir = Path(output_dir)
            
            if not docx_path.exists():
                print(f"Error: Input file {docx_path} does not exist")
                return False
            
            if not docx_path.suffix.lower() == '.docx':
                print(f"Error: Input file must have .docx extension")
                return False
            
            # Create output directory if it doesn't exist
            output_dir.mkdir(parents=True, exist_ok=True)
            
            # Extract the .docx file (which is a ZIP archive)
            with zipfile.ZipFile(docx_path, 'r') as zip_ref:
                zip_ref.extractall(output_dir)
            
            # Pretty print XML files for better readability
            self._prettify_xml_files(output_dir)
            
            print(f"Successfully converted {docx_path} to OpenXML in {output_dir}")
            return True
            
        except Exception as e:
            print(f"Error converting DOCX to OpenXML: {str(e)}")
            return False
    
    def openxml_to_docx(self, openxml_dir, docx_path):
        """
        Convert OpenXML directory structure back to a .docx file.
        
        Args:
            openxml_dir (str): Directory containing OpenXML files
            docx_path (str): Path for the output .docx file
        
        Returns:
            bool: True if successful, False otherwise
        """
        try:
            openxml_dir = Path(openxml_dir)
            docx_path = Path(docx_path)
            
            if not openxml_dir.exists():
                print(f"Error: OpenXML directory {openxml_dir} does not exist")
                return False
            
            # Ensure output has .docx extension
            if not docx_path.suffix.lower() == '.docx':
                docx_path = docx_path.with_suffix('.docx')
            
            # Create parent directory if it doesn't exist
            docx_path.parent.mkdir(parents=True, exist_ok=True)
            
            # Create ZIP archive with .docx extension
            with zipfile.ZipFile(docx_path, 'w', zipfile.ZIP_DEFLATED) as zip_ref:
                for file_path in openxml_dir.rglob('*'):
                    if file_path.is_file():
                        # Calculate relative path to maintain directory structure
                        arc_name = file_path.relative_to(openxml_dir)
                        zip_ref.write(file_path, arc_name)
            
            print(f"Successfully converted OpenXML from {openxml_dir} to {docx_path}")
            return True
            
        except Exception as e:
            print(f"Error converting OpenXML to DOCX: {str(e)}")
            return False
    
    def _prettify_xml_files(self, directory):
        """
        Pretty print XML files in the directory for better readability.
        
        Args:
            directory (Path): Directory containing XML files
        """
        xml_extensions = {'.xml', '.rels'}
        
        for file_path in directory.rglob('*'):
            if file_path.is_file() and file_path.suffix.lower() in xml_extensions:
                try:
                    # Read and parse XML
                    with open(file_path, 'r', encoding='utf-8') as f:
                        content = f.read()
                    
                    # Pretty print
                    parsed = minidom.parseString(content)
                    pretty_xml = parsed.toprettyxml(indent='  ')
                    
                    # Remove empty lines and write back
                    lines = [line for line in pretty_xml.split('\n') if line.strip()]
                    with open(file_path, 'w', encoding='utf-8') as f:
                        f.write('\n'.join(lines))
                        
                except Exception as e:
                    # If XML parsing fails, leave file as is
                    print(f"Warning: Could not prettify {file_path}: {str(e)}")
    
    def validate_docx_structure(self, path):
        """
        Validate that a .docx file has the correct structure.
        
        Args:
            path (str): Path to the .docx file
        
        Returns:
            bool: True if valid, False otherwise
        """
        try:
            with zipfile.ZipFile(path, 'r') as zip_ref:
                file_list = zip_ref.namelist()
                
                # Check for essential files
                required_files = ['[Content_Types].xml', '_rels/.rels']
                for req_file in required_files:
                    if req_file not in file_list:
                        print(f"Error: Missing required file: {req_file}")
                        return False
                
                # Check for word document structure
                has_document = any('word/document.xml' in f for f in file_list)
                if not has_document:
                    print("Error: Missing word/document.xml")
                    return False
                
                return True
                
        except Exception as e:
            print(f"Error validating DOCX structure: {str(e)}")
            return False


def main():
    """Main function to handle command line arguments and execute conversions."""
    parser = argparse.ArgumentParser(description='Convert between DOCX and OpenXML formats')
    parser.add_argument('operation', choices=['to-openxml', 'to-docx'], 
                       help='Operation to perform')
    parser.add_argument('input', help='Input file or directory')
    parser.add_argument('output', help='Output file or directory')
    parser.add_argument('--validate', action='store_true', 
                       help='Validate file structure before conversion')
    
    args = parser.parse_args()
    
    converter = DOCXOpenXMLConverter()
    
    if args.operation == 'to-openxml':
        if args.validate:
            if not converter.validate_docx_structure(args.input):
                print("Validation failed. Aborting conversion.")
                return
        
        success = converter.docx_to_openxml(args.input, args.output)
        
    elif args.operation == 'to-docx':
        success = converter.openxml_to_docx(args.input, args.output)
    
    if success:
        print("Conversion completed successfully!")
    else:
        print("Conversion failed!")


# Example usage functions
def example_usage():
    """Example of how to use the converter programmatically."""
    converter = DOCXOpenXMLConverter()
    
    # Convert DOCX to OpenXML
    docx_file = "example.docx"
    openxml_dir = "example_openxml"
    
    if converter.docx_to_openxml(docx_file, openxml_dir):
        print("DOCX to OpenXML conversion successful")
        
        # Convert back to DOCX
        new_docx_file = "example_converted.docx"
        if converter.openxml_to_docx(openxml_dir, new_docx_file):
            print("OpenXML to DOCX conversion successful")


if __name__ == "__main__":
    main()