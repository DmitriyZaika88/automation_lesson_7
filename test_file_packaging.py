import pytest
import os
import zipfile
import PyPDF2
import openpyxl


@pytest.fixture
def zip_file():
    file_names = [os.path.abspath('pdf_example.pdf'), os.path.abspath('xlsx_example.xlsx'), os.path.abspath('csv_example.csv')]

    zip_file = zipfile.ZipFile('resources/archive.zip', 'w')
    for file_name in file_names:
        if os.path.exists(file_name):
            zip_file.write(file_name, os.path.basename(file_name))
        else:
            print(f"Файл {file_name} не найден!")
    zip_file.close()


def test_read_pdf(zip_file):
    with zipfile.ZipFile('resources/archive.zip', 'r') as zip_file:
        with zip_file.open('pdf_example.pdf', 'r') as pdf_file:
            pdf_reader = PyPDF2.PdfReader(pdf_file)
            assert len(pdf_reader.pages) == 1, "PDF file should have 1 page."


def test_read_csv(zip_file):
    with zipfile.ZipFile('resources/archive.zip', 'r') as zip_file:
        with zip_file.open('csv_example.csv', 'r') as csv_file:
            count_row = 0
            for _ in csv_file:
                count_row += 1
            assert 44 == count_row


def test_read_xlsx(zip_file):
    with zipfile.ZipFile('resources/archive.zip', 'r') as zip_file:
        with zip_file.open('xlsx_example.xlsx', 'r') as xlsx_file:
            book = openpyxl.load_workbook(xlsx_file)
            sheet = book.active
            print(sheet.cell(row=2, column=2).value)
            assert 'Title' in sheet['A1'].value
            assert 'John Doe' in sheet['G2'].value
