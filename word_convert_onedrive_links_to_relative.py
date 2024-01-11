import os
import docx
from urllib.parse import urljoin

from docx import Document
def convert_links_to_relative(folder_path, file_path, onedrive_url):
    updated_count = 0
    onedrive_folder_url = urljoin(onedrive_url + "/", os.path.relpath(folder_path, os.getenv("OneDrive")).replace("\\", "/"))
    relative_path = os.path.relpath(file_path, folder_path)
    depth = len(os.path.normpath(relative_path).split(os.path.sep)) - 1
    doc = Document(file_path)

    # Iterate through paragraphs
    for paragraph in doc.paragraphs:
        if paragraph.hyperlinks:
            for hyperlink in paragraph.hyperlinks:
                if onedrive_folder_url in hyperlink.address:
                    # Update the address with the new URL
                    part = paragraph.part
                    r_id = part.relate_to(hyperlink.address.replace(onedrive_folder_url, "..\\" * depth if depth > 0 else ""), docx.opc.constants.RELATIONSHIP_TYPE.HYPERLINK, is_external=True)
                    new_hyperlink = docx.oxml.shared.OxmlElement('w:hyperlink')
                    new_hyperlink.set(docx.oxml.shared.qn('r:id'), r_id, )
                    new_run = docx.text.run.Run(docx.oxml.shared.OxmlElement('w:r'), paragraph)
                    new_run.text = hyperlink.text
                    new_run.style = get_or_create_hyperlink_style(part.document)
                    new_hyperlink.append(new_run._element)
                    index = next((i for i, run in enumerate(paragraph._p) if run.text == hyperlink.text), None)
                    paragraph._p[index] = new_hyperlink
                    updated_count += 1

    # Save the updated document
    if(updated_count > 0):
        print("updated:", str(updated_count), "links in", os.path.basename(file_path))
        doc.save(file_path.replace(".docx", "_updated.docx"))

def get_or_create_hyperlink_style(d):
    """If this document had no hyperlinks so far, the builtin
       Hyperlink style will likely be missing and we need to add it.
       There's no predefined value, different Word versions
       define it differently.
       This version is how Word 2019 defines it in the
       default theme, excluding a theme reference.
    """
    if "Hyperlink" not in d.styles:
        if "Default Character Font" not in d.styles:
            ds = d.styles.add_style("Default Character Font",
                                    docx.enum.style.WD_STYLE_TYPE.CHARACTER,
                                    True)
            ds.element.set(docx.oxml.shared.qn('w:default'), "1")
            ds.priority = 1
            ds.hidden = True
            ds.unhide_when_used = True
            del ds
        hs = d.styles.add_style("Hyperlink",
                                docx.enum.style.WD_STYLE_TYPE.CHARACTER,
                                True)
        hs.base_style = d.styles["Default Character Font"]
        hs.unhide_when_used = True
        hs.font.color.rgb = docx.shared.RGBColor(0x05, 0x63, 0xC1)
        hs.font.underline = True
        del hs

    return "Hyperlink"


def find_word_files_recursively(folder_path):
    word_files = []
    for root, dirs, files in os.walk(folder_path):
        for file in files:
            if file.endswith('.docx'):
                word_files.append(os.path.join(folder_path, root, file))
    return word_files


def main():
    folder_path = input("Path to folder: ")
    while not os.path.isdir(folder_path):
        print(f'Error: The provided path "{folder_path}" is not a valid directory.')
        folder_path = input("Path to folder: ")
    onedrive_url = input("Onedrive base url: ")
    folder_path = os.path.abspath(folder_path)
    files = find_word_files_recursively(folder_path)
    for file_path in files:
        convert_links_to_relative(folder_path, file_path, onedrive_url)

if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        print(e)
    input("Press Enter to exit...")
