from utils.read_excel import Excel_file
from utils.read_pdfs import PDF_Utils

import argparse


def main():
    parser = argparse.ArgumentParser(
        description="Processes the files and creates a manual for you"
    )

    parser.add_argument(
        "-i", "--input", type=str, help="Input PDF folder path", required=True
    )
    parser.add_argument(
        "-o", "--output", type=str, help="Output path for the PDF file", required=False
    )

    parser.add_argument(
        "-e" "--excel",
        type=str,
        help="Excel file path for creating indexes",
        required=True,
    )

    args = parser.parse_args()

    pdf_utils = PDF_Utils(
        args.input, args.output, template_path="./pdfs/TemplateWordA4v01.pdf"
    )

    # reads the pdf files in the provided folder
    pdf_dict = pdf_utils.read_pdfs()
    output = args.output or "output.pdf"

    pdf_utils.merge_files(pdf_dict, output)

    # up until this point, the pdf file is created, merging all the files in the folder
    # and that's it, from now on, we only create the index and necessary bookmarks
    # everything from now on is only creating indexes, editing should be done
    # before creating the index!!

    # creates bookmarks
    excel_reader = Excel_file("./10008361744.xlsm")
    index_items = excel_reader.create_index(True)
    pdf_utils.create_index_and_merge(
        args.output, index_items
    )  # pdf with bookmarks only!!


if __name__ == "__main__":
    main()
