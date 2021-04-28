import sys
import os
import zipfile
import shutil
from pathlib import Path


def replace_path_suffix(path, _suffix_length, _new_suffix):
    """
    Replace the suffix of a path with new suffix
    :param path:            original path
    :param _suffix_length:  the number of chars to be deleted
                            from the end of path
    :param new_suffix:      the new suffix to be added to path
    :return:                new path where the last _suffix_length
                            chars are replaced with _new_suffix
    """
    new_path = path[0:-_suffix_length] + _new_suffix
    return new_path


def docx_to_zip(_docx_path):
    """
    Copy a docx file (to preserve the original file) and change
    the duplicated docx file to zip file
    :param _docx_path: the path of the docx file to be duplicated
    :return:           the new zip file path
    """
    _zip_path = replace_path_suffix(_docx_path, _suffix_length=4, _new_suffix="zip")
    shutil.copyfile(_docx_path, _zip_path)
    return _zip_path


def extract_zip_file(_zip_path):
    """
    Extract zip file to a directory with the same name as the zip file
    :param _zip_path: the zip file path to be extracted
    :return:          the extraction directory path
    """
    with zipfile.ZipFile(_zip_path, 'r') as zip_ref:
        _extraction_directory = replace_path_suffix(_zip_path, _suffix_length=4, _new_suffix="/")
        # _extraction_directory = _zip_path[0:-4] + '/'
        zip_ref.extractall(_extraction_directory)
        return _extraction_directory


def equations_to_text(_document_xml_path, _out_xml_path):
    """
    Generate new xml file identical to a given document.xml file
    such that all the equations are converted to text equations
    :param _document_xml_path: the original document.xml file path
    :param _out_xml_path:      the new xml file path
    """
    with open(_document_xml_path, "rt", encoding="utf8") as fin:
        with open(_out_xml_path, "wt", encoding="utf8") as fout:
            for line in fin:
                fout.write(line.replace('<m:r><w:rPr>', '<m:r><m:rPr><m:nor /></m:rPr><w:rPr>'))


def create_zip_file(_zip_path, directory):
    """
    Create a zip file from all the files in a directory
    :param _zip_path: the new zip path
    :param directory: the directory to be archived
    """
    with zipfile.ZipFile(_zip_path, 'w') as zip_obj:
        for folder_name, sub_folders, filenames in os.walk(directory):
            for filename in filenames:
                file_path = os.path.join(folder_name, filename)
                p = Path(file_path)
                # add a file to the zip file in the relative directory path
                zip_obj.write(file_path, p.relative_to(directory))


def zip_to_docx(_zip_path):
    """
    Change zip file to docx file
    :param _zip_path: the zip file to be changed path
    :return:          the docx file path
    """
    _docx_path = replace_path_suffix(_zip_path, _suffix_length=4, _new_suffix="_f.docx")
    os.rename(_zip_path, _docx_path)
    return _docx_path


# Script Purpose
# --------------
# Given a docx file containing equations
# Generate new docx file where all the equations are converted to text
# The new docx file is named as the original docx file name + '_f'
# ==============
# Script Logic
# --------------
#   1. Get the docx file path as parameter
#   2. Copy the docx file and convert the duplicated file to zip file
#   3. Extract the zip file
#   4. Change all the equations to text in document.xml
#   5. Archive the new files to zip file
#   6. Convert the zip file to docx file

ORIGINAL_DOCX_PATH_INDEX = 1

#   1
original_docx_path = sys.argv[ORIGINAL_DOCX_PATH_INDEX]
docx_path = os.path.normpath(original_docx_path)
#   2
zip_path = docx_to_zip(docx_path)
#   3
extraction_directory = extract_zip_file(zip_path)
#   4
document_xml_path = extraction_directory + 'word/document.xml'
tmp_document_xml_path = extraction_directory + 'word/document_tmp.xml'
equations_to_text(document_xml_path, tmp_document_xml_path)
os.replace(tmp_document_xml_path, document_xml_path)
#   5
create_zip_file(zip_path, extraction_directory)
shutil.rmtree(extraction_directory)
#   6
result_docx_path = zip_to_docx(zip_path)
