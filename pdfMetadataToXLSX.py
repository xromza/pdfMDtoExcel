"""
This script parses .pdf files and stores its metadata in .xlsx file
create_metadata function can handle multiple .pdf files, but spaces between letters are not allowed and causes some errors
=================================
by xromza
"""
def install_requirements() -> None:
    """
    Installing pypdf using pip
    """
    try:
        from pip import main as pipmain
    except ImportError: # Backward compatibility
        from pip._internal import main as pipmain
    pipmain(['install', "-q", "pypdf", "pandas", "openpyxl"]) # Installing pypdf and pandas

def create_metadata(namesOfFiles = '') -> None: # creates a func to use externally
    while namesOfFiles == '' or namesOfFiles == " ": # If function called without giving parameters
        namesOfFiles = input("Enter name of PDF files in current directory: ")
    import os
    try:
        from pypdf import PdfReader # Imports pyPDF library
        import pandas as pd # Imports pandas module
        import openpyxl # Imports library for pandas
    except ImportError: # Install required library if it missing
        install_requirements()
        from pypdf import PdfReader # Imports pyPDF library again
        import pandas as pd # Imports pandas library again
    for file in namesOfFiles.split(','): # Multiple .pdf
        file = ''.join(['' if (x == " ") else x for x in file]) # Delete spaces
        if ''.join(file)[-4::] != '.pdf': # If name of file doesn't content ".pdf"
            file += '.pdf' # add ".pdf" to user input

        print("\nParsing metadata from " + file)
        
        dir_path = os.path.dirname(os.path.realpath(__file__)) # Full address to current directory
        try:
            readPDF = PdfReader(file) # reads pdf file
        except FileNotFoundError:
            print('File not founded, skipping. (Spaces are not allowed)') # if file with entered name is not found, skips that file
            continue
        metadataFileName = file[:len(file)-4:] + '_metadata.xlsx' # name for the new metadata .xlsx file

        pd.DataFrame.from_dict([dict(readPDF.metadata)]).to_excel(metadataFileName) # Writing Excel table

        print("DONE\nFile stored in " + dir_path + "\\" + metadataFileName)
if __name__ == "__main__": # main script for launching directly
    create_metadata(input("Enter name of PDF files in current directory: "))
