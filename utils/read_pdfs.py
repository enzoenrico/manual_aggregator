import os
from typing import Dict, List
from pypdf import PageObject, PdfReader, PdfWriter
import pypdf
from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas
from reportlab.lib.colors import black, gray
import io


class PDF_Utils:
    def __init__(self, folder_path: str, pdf_path: str, template_path: str) -> None:
        self.folder_path = folder_path
        self.template_path = template_path
        self.pdf_path = pdf_path

    def read_pdfs(self) -> Dict[str, PdfReader]:
        files: Dict[str, PdfReader] = {}
        for filename in os.listdir(self.folder_path):
            if "TemplateWord" not in filename:
                pdf_path = os.path.join(self.folder_path, filename)
            if os.path.isfile(pdf_path) and pdf_path.endswith(".pdf"):
                files[pdf_path] = PdfReader(pdf_path)
        return files

    def merge_files(self, file_dict: Dict[str, PdfReader], output_path: str) -> None:
        merger = PdfWriter()
        for file in file_dict.keys():
            merger.append(file)
        with open(output_path, "wb") as f:
            merger.write(f)
        merger.close()

    def make_index_page(self, index: Dict[int, List[Dict]]) -> PageObject:
        template_page = self.template_path

        cv = canvas.Canvas(template_page, pagesize=letter)
        cv.setFont("Helvetica", 12)

        for idx, (key, val) in enumerate(index.items()):
            cv.setFillColor(black)
            cv.setFontSize(11)

            next_topic_jump = len(val) * 20 + 10

            cv.drawString(100, 750 - next_topic_jump, f"Index {key}: {len(val)} topics")
            for indexes, items in enumerate(val):
                cv.setFillColor(gray)
                cv.setFontSize(10)
                cv.drawString(
                    120,
                    740 - indexes * 20 - (idx + 1) * 20,
                    f"{items['Descrição']}",
                )

        cv.save()
        packet = io.BytesIO()
        packet.write(open(template_page, "rb").read())
        packet.seek(0)
        new_pdf = PdfReader(packet)
        return new_pdf.pages[0]

    def create_index_and_merge(
        self, merged_file_path: str, index: Dict[int, List[Dict]]
    ) -> None:
        merged_file = PdfReader(merged_file_path)
        merged_with_index = PdfWriter()

        # TODO
        # change for adding outlines to the actual pages
        # need to find the pages in the document

        index_page = self.make_index_page(index)
        merged_with_index.add_page(index_page)


        # count the pages on each added document, but how am i gonna track
        # what docs are being added?
        # -e

        section_indexes = [
            "Section: " + str(title) for i, title in enumerate(index.keys())
        ]

        # section_sub_indexes = []

        for i, v in enumerate(index.values()):
            parent_outline = merged_with_index.add_outline_item(
                title=section_indexes[i], page_number=i + 1, bold=True
            )
            for _, values in enumerate(v):
                # index_titles = {
                #     "parent": section_indexes[i],
                #     "value": values["Descrição"],
                # }

                merged_with_index.add_outline_item(
                    title=values["Descrição"],
                    #change here to right addresses
                    page_number=i + 1,
                    parent=parent_outline,
                )
        for _, v in enumerate(PdfReader(merged_file_path).pages):
            merged_with_index.add_page(v)
        # section_sub_indexes.append(index_titles)

        # pprint.pprint(section_sub_indexes)

        # merged_with_index.add_page(page)
        # parent_outline = merged_with_index.add_outline_item(
        #     title=index_titles[1],
        #     page_number=idx + 1,
        #     bold=True,
        # )

        merged_with_index.add_page(merged_file.pages[0])

        with open(merged_file_path, "wb") as f:
            merged_with_index.write(f)
            # template = open('../pdfs/TemplateWordA4v01.pdf', "wb")
            # merged_with_index.write(template)

        merged_with_index.close()
