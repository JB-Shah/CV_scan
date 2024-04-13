from django.shortcuts import render
from django.http import HttpResponse
from django.core.files.storage import FileSystemStorage
import os
import re
import pandas as pd
from docx import Document
from PyPDF2 import PdfReader
import openpyxl

def extract_info_from_docx(docx_file):
    text = ""
    doc = Document(docx_file)
    for para in doc.paragraphs:
        text += para.text + "\n"
    return text

def extract_info_from_pdf(pdf_file):
    text = ""
    with open(pdf_file, "rb") as file:
        reader = PdfReader(file)
        for page_num in range(len(reader.pages)):
            text += reader.pages[page_num].extract_text()
    return text

def extract_email_addresses(text):
    email_pattern = r'\b[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Z|a-z]{2,}\b'
    emails = re.findall(email_pattern, text)
    return emails[0] if emails else ""

def extract_phone_numbers(text):
    phone_pattern = r'\b(?:\d{5}[-.\s]??\d{5}|\d{4}[-.\s]??\d{5}|\d{3}[-.\s]??\d{8}|\d{10})\b'
    phones = re.findall(phone_pattern, text)
    return phones[0] if phones else ""

def convert_to_excel(data_list, excel_file_path):
    df = pd.concat(data_list)
    df.to_excel(excel_file_path, index=False)

def upload_file(request):
    if request.method == 'POST' and request.FILES.getlist('document'):
        data_list = []
        for uploaded_file in request.FILES.getlist('document'):
            fs = FileSystemStorage()
            filename = fs.save(uploaded_file.name, uploaded_file)
            uploaded_file_path = os.path.join(fs.location, filename)

            if filename.endswith('.docx'):
                text = extract_info_from_docx(uploaded_file_path)
            elif filename.endswith('.pdf'):
                text = extract_info_from_pdf(uploaded_file_path)
            else:
                return HttpResponse(f"Unsupported file format: {filename}. Only DOCX and PDF files are supported.")

            email = extract_email_addresses(text)
            phone = extract_phone_numbers(text)

            data = {
                "Email": [email],
                "Phone Number": [phone],
                "Text": [text]
            }
            df = pd.DataFrame(data)
            data_list.append(df)

        excel_filename = "cv_data.xlsx"
        excel_file_path = os.path.join(fs.location, excel_filename)
        convert_to_excel(data_list, excel_file_path)

        with open(excel_file_path, 'rb') as excel_file:
            response = HttpResponse(excel_file.read(), content_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
            response['Content-Disposition'] = f'attachment; filename="{excel_filename}"'
            return response

    return render(request, 'core/upload_file.html')