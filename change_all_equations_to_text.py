import sys
import os
import zipfile
import shutil
from pathlib import Path


def docx_to_zip(docx_file_path):
    zip_file_path = docx_file_path[0:-4] + "zip"
    new_docx_file_path = docx_file_path + '_copy'
    shutil.copyfile(docx_file_path, new_docx_file_path)
    os.rename(new_docx_file_path, zip_file_path)
    return zip_file_path


def extract_zip_file(zip_file_path):
    with zipfile.ZipFile(zip_file_path, 'r') as zip_ref:
        extract_directory = zip_file_path[0:-4] + '/'
        zip_ref.extractall(extract_directory)
        return extract_directory


def equations_to_text(document_xml_file_path, out_xml_file_path):
    with open(document_xml_file_path, "rt", encoding="utf8") as fin:
        with open(out_xml_file_path, "wt", encoding="utf8") as fout:
            for line in fin:
                fout.write(line.replace('<m:r><w:rPr>', '<m:r><m:rPr><m:nor /></m:rPr><w:rPr>'))


def zip_file(zip_file_path, directory):
    with zipfile.ZipFile(zip_file_path, 'w') as zipObj:
        for folder_name, sub_folders, filenames in os.walk(directory):
            for filename in filenames:
                file_path = os.path.join(folder_name, filename)
                p = Path(file_path)
                zipObj.write(file_path, p.relative_to(directory))


def zip_to_docx(zip_file_path):
    docx_file_path = zip_file_path[0:-4] + "_f.docx"
    os.rename(zip_file_path, docx_file_path)
    # copyfile(docx_file_path, zip_file_path)
    return docx_file_path



# print(sys.argv[1])
orig_docx_file_path1 = sys.argv[1]
docx_file_path1 = os.path.normpath(orig_docx_file_path1)
# docx_file_path1 = "C:/Users/mitmi/Desktop/Equations Project/Test/equations.docx"
zip_file_path1 = docx_to_zip(docx_file_path1)
extract_directory1 = extract_zip_file(zip_file_path1)
document_xml_file_path1 = extract_directory1 + 'word/document.xml'
tmp_document_xml_file_path1 = extract_directory1 + 'word/document1.xml'
equations_to_text(document_xml_file_path1, tmp_document_xml_file_path1)
os.replace(tmp_document_xml_file_path1, document_xml_file_path1)
zip_file(zip_file_path1, extract_directory1)
new_docx_file_path1 = zip_to_docx(zip_file_path1)
shutil.rmtree(extract_directory1)


# xml_file_path1 = "C:/Users/mitmi/Desktop/Equations Project/Test/equations/word/document.xml"
# new_xml_file_path1 = "C:/Users/mitmi/Desktop/Equations Project/Test/equations/word/document1.xml"
# equations_to_text(xml_file_path1, new_xml_file_path1)
