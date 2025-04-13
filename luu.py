import os
import pdfplumber
import pandas as pd
import re
import pytesseract
from pdf2image import convert_from_path
from PIL import Image

def extract_text_from_pdf(pdf_path):
    """Trích xuất văn bản từ PDF, hỗ trợ OCR nếu cần."""
    text = ""
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            extracted = page.extract_text()
            if extracted:
                text += extracted + "\n\n"

    # Nếu không có văn bản, thực hiện OCR trên ảnh
    if not text.strip():
        images = convert_from_path(pdf_path)
        for img in images:
            text += pytesseract.image_to_string(img, lang='eng+vie') + "\n\n"

    text = text.strip()
    word_count = len(text.split())  # Đếm số từ
    print(f"✅ Trích xuất xong từ PDF: {pdf_path} - Tổng số từ: {word_count}")

    return text


def extract_bold_titles(pdf_path):
    """Trích xuất tiêu đề có số thứ tự từ PDF, hỗ trợ cả in đậm và không in đậm."""
    titles = []
    pattern = re.compile(r'^\d+(\.\d+)*\s+.+')  # Cải tiến regex

    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            text = page.extract_text()
            if text:
                for line in text.split("\n"):
                    line = line.strip()
                    if pattern.match(line):
                        titles.append(line)

    return titles


def split_text_into_sections(text, titles):
    """Tách văn bản thành các chương và mục dựa trên tiêu đề có số thứ tự."""
    sections = {}
    current_chapter = "Giới thiệu"
    current_title = ""
    sections[current_chapter] = {}

    title_pattern = re.compile(r'^\d+(\.\d+)*\s+.+')

    for line in text.split("\n"):
        line = line.strip()

        # Nhận diện chương mới (VD: "CHƯƠNG 1: Giới thiệu")
        if re.match(r'^CHƯƠNG \d+:?\s+.+', line):
            current_chapter = line
            sections.setdefault(current_chapter, {})
            current_title = ""

        # Nhận diện tiêu đề có số thứ tự
        elif line in titles or title_pattern.match(line):
            current_title = line
            sections[current_chapter].setdefault(current_title, "")

        # Nội dung đoạn văn bản
        elif current_title:
            sections[current_chapter][current_title] += line + " "

    return sections


def split_paragraphs(text):
    """Tách văn bản thành đoạn văn."""
    paragraphs = re.split(r'(?<=[.!?])\s+', text)
    return [p.strip() for p in paragraphs if p.strip()]

def merge_paragraphs(paragraphs):
    """Ghép đoạn văn có cùng chủ đề thành đoạn dài 300-700 từ."""
    merged_paragraphs = []
    buffer = ""
    for para in paragraphs:
        if len(buffer.split()) + len(para.split()) <= 700:
            buffer += (" " if buffer else "") + para
        else:
            if len(buffer.split()) >= 300:
                merged_paragraphs.append(buffer.strip())
            buffer = para
    
    if 300 <= len(buffer.split()) <= 700:
        merged_paragraphs.append(buffer.strip())
    
    return merged_paragraphs

def generate_data(pdf_path, sections):
    """Tạo dữ liệu gồm tên file, chương, tiêu đề mục và nội dung."""
    extracted_data = []
    file_name = os.path.basename(pdf_path)  # Lấy tên file PDF
    for chapter, titles in sections.items():
        for title, text in titles.items():
            paragraphs = split_paragraphs(text)
            merged_paragraphs = merge_paragraphs(paragraphs)
            for para in merged_paragraphs:
                extracted_data.append((file_name, chapter, title, len(para.split()), para))
    return extracted_data

def save_to_excel(data, output_path):
    """Lưu dữ liệu vào file Excel."""
    df = pd.DataFrame(data, columns=["Tên File", "Chương", "Tiêu đề Mục", "Độ dài", "Bản Gốc"])
    df.to_excel(output_path, index=False)
    print(f"✅ Dữ liệu đã được lưu vào: {os.path.abspath(output_path)}")

def process_pdf(pdf_path):
    """Xử lý từng file PDF và trích xuất dữ liệu."""
    if not os.path.exists(pdf_path):
        print(f"⚠️ File không tồn tại: {pdf_path}")
        return []
    
    text = extract_text_from_pdf(pdf_path)
    if not text:
        print(f"⚠️ Không tìm thấy nội dung trong PDF: {pdf_path}")
        return []
    
    titles = extract_bold_titles(pdf_path)
    sections = split_text_into_sections(text, titles)
    return generate_data(pdf_path, sections)

# Danh sách file PDF cần xử lý
pdf_files = [
    # r"C:\Tailieuht\TA\Giaotrinh\Đã Làm\giaotrinhlaptrinhc.pdf",
    r"C:\Tailieuht\TA\Giaotrinh\Đã Làm\lap_trinh_mang.pdf",
    # r"C:\Tailieuht\TA\Giaotrinh\GT-tri-tue-nhan-tao-1.pdf",
    # r"C:\Tailieuht\TA\Giaotrinh\gtantoanvabaomatthongtin.pdf",
]

output_excel = r"C:\Tailieuht\TA\Giaotrinh\Đã Làm\t0.xlsx"

# Lưu dữ liệu từ tất cả các file PDF
all_data = []
for pdf_path in pdf_files:
    all_data.extend(process_pdf(pdf_path))

if all_data:
    save_to_excel(all_data, output_excel)
else:
    print("⚠️ Không có dữ liệu nào được trích xuất!")
