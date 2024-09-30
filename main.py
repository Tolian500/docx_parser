from docx import Document
from lxml import etree
import hashlib


# Function to generate a class name based on font, size, and styles combination
def generate_class_name(font, size, is_bold, is_italic, is_underlined):
    style_suffix = ""
    if is_bold:
        style_suffix += "bold"
    if is_italic:
        style_suffix += "italic"
    if is_underlined:
        style_suffix += "underline"

    return f'font_{hashlib.md5(f"{font}_{size}_{style_suffix}".encode()).hexdigest()[:8]}'


# Function to convert the .docx file to an XML format with fonts, sizes, alignment, and indentation defined as styles
def docx_to_xml_with_styles(docx_file, xml_output):
    # Read the DOCX file
    doc = Document(docx_file)

    # Get default page margins from the first section
    section = doc.sections[0]
    top_margin = section.top_margin.pt
    bottom_margin = section.bottom_margin.pt
    left_margin = section.left_margin.pt
    right_margin = section.right_margin.pt

    # Create the root element for the XML structure
    root = etree.Element('html', xmlns="http://www.w3.org/1999/xhtml", xmlns_th="http://www.thymeleaf.org")

    # Add <head> and <style> sections to the XML
    head = etree.SubElement(root, 'head')
    meta = etree.SubElement(head, 'meta', attrib={
        'http-equiv': 'Content-Type',
        'content': 'text/html; charset=UTF-8'
    })

    # Add general body style to the XML, using extracted margins
    style = etree.SubElement(head, 'style')
    style.text = f"""
        body {{
            font-size: 11px;
            text-align: justify;
            margin: {top_margin}pt {right_margin}pt {bottom_margin}pt {left_margin}pt;  /* Set default page margin */
        }}
    """

    # Dictionary to hold all font and size combinations
    global_styles = {}

    # Add <body> section
    body = etree.SubElement(root, 'body')

    # Parse paragraphs from the DOCX file and add them to the XML structure
    for para in doc.paragraphs:
        # Determine the alignment and set the appropriate class
        alignment_class = ""
        if para.alignment == 1:  # Center
            alignment_class = "center"
        elif para.alignment == 2:  # Right
            alignment_class = "right"
        elif para.alignment == 3:  # Justified
            alignment_class = "justify"
        else:  # Left (default)
            alignment_class = "left"

        # Get indentation properties (left indentation)
        left_indent = para.paragraph_format.left_indent.pt if para.paragraph_format.left_indent else 0
        # Create <p> element for each paragraph with alignment class and indentation
        p = etree.SubElement(body, 'p', attrib={
            'class': f'common-style {alignment_class}',
            'style': f'margin-left: {left_indent}pt;'
        })

        # Extract font styles for each run in the paragraph
        for run in para.runs:
            font_name = run.font.name if run.font.name else "default"
            font_size = run.font.size.pt if run.font.size else 11  # Default to 11pt if no size
            is_bold = run.bold
            is_italic = run.italic
            is_underlined = run.underline

            # Generate a class name for the font/size/style combination
            class_name = generate_class_name(font_name, font_size, is_bold, is_italic, is_underlined)

            # Handle tab characters explicitly
            run_text = run.text
            if run_text:
                # Replace tab characters with spaces (you can adjust the number of spaces as needed)
                run_text = run_text.replace('\t', '    ')  # Using 4 spaces for each tab

                # Only create a span if the style differs from the common style
                if font_name != "default" or font_size != 11 or is_bold or is_italic or is_underlined:
                    # Store the class name in the global styles dictionary if it's not already there
                    if class_name not in global_styles:
                        global_styles[class_name] = {
                            'font': font_name,
                            'size': font_size,
                            'is_bold': is_bold,
                            'is_italic': is_italic,
                            'is_underlined': is_underlined,
                        }

                    # Create a <span> element with the corresponding class
                    span = etree.SubElement(p, 'span', attrib={
                        'class': class_name
                    })
                    span.text = run_text
                else:
                    # If the style matches the common style, directly add text to the paragraph
                    if p.text:
                        p.text += run_text
                    else:
                        p.text = run_text

    # Append the collected styles into the <style> tag in the <head> section
    for class_name, style_info in global_styles.items():
        font_family = style_info['font']
        font_size = style_info['size']
        font_weight = 'bold' if style_info['is_bold'] else 'normal'
        font_style = 'italic' if style_info['is_italic'] else 'normal'
        text_decoration = 'underline' if style_info['is_underlined'] else 'none'

        style.text += f'.{class_name} {{ font-family: "{font_family}"; font-size: {font_size}pt; font-weight: {font_weight}; font-style: {font_style}; text-decoration: {text_decoration}; }}\n'

    # Add styles for text alignment
    style.text += """
        p.left {
            text-align: left;
        }
        p.center {
            text-align: center;
        }
        p.right {
            text-align: right;
        }
        p.justify {
            text-align: justify;
        }
    """

    # Write the XML to a file
    tree = etree.ElementTree(root)
    tree.write(xml_output, pretty_print=True, xml_declaration=True, encoding="UTF-8")


# Paths to the .docx file and output XML file
path = "20230531.docx"
output_path = 'text_with_styles.xml'

# Convert DOCX to XML with fonts, sizes, alignment, and indentation as styles
docx_to_xml_with_styles(path, output_path)

print(f"XML content generated at {output_path}")
