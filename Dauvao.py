import os
import pdfplumber
import pandas as pd
import re
import pytesseract
from pdf2image import convert_from_path

def extract_text_from_pdf(pdf_path):
    """Trích xuất văn bản từ PDF, hỗ trợ OCR nếu cần."""
    text = ""
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            extracted = page.extract_text()
            if extracted:
                text += extracted + "\n\n"

    if not text.strip():  # Nếu không trích xuất được, dùng OCR
        images = convert_from_path(pdf_path)
        custom_config = r'--oem 3 --psm 6'
        for img in images:
            text += pytesseract.image_to_string(img, lang='vie', config=custom_config) + "\n\n"

    text = clean_vietnamese_text(text)
    return text.strip()

def clean_vietnamese_text(text):
    """Sửa lỗi ký tự bị trích xuất sai từ PDF."""
    replacements = {
        "l{": "là",
        "Đ}": "Đây",
        "c|c": "các",
        "t}m": "tìm",
        "{": "à",
        "}": "â",
        "|": "á",
        "m{ng": "mạng",
        "ph}ng": "phòng",
        "Ö": "ệ",
        "−¬": "ươ",
        "Ƣ": "ư",
        "~": "ã",
        "ƣ": "ư",
    }

    for wrong, correct in replacements.items():
        text = text.replace(wrong, correct)

    return text

def extract_titles(text):
    """Trích xuất tiêu đề từ văn bản."""
    titles = []
    pattern1 = re.compile(r'^(?!\d{2,}\.)(\d+(\.\d+)+)\s+([A-ZĐÁÀẢÃẠĂẮẰẲẴẶÂẤẦẨẪẬÊẾỀỂỄỆÔỐỒỔỖỘƠỚỜỞỠỢƯỨỪỬỮỰ][\w\s,()-]*)$')
    pattern2 = re.compile(r'^(\d+(\.\d+)*\.)\s*([A-ZĐÁÀẢÃẠĂẮẰẲẴẶÂẤẦẨẪẬÊẾỀỂỄỆÔỐỒỔỖỘƠỚỜỞỠỢƯỨỪỬỮỰa-z][\w\s,()-]*)$', re.UNICODE)
    pattern3 = re.compile(r'^(\d+)\.\s+([A-ZĐÁÀẢÃẠĂẮẰẲẴẶÂẤẦẨẪẬÊẾỀỂỄỆÔỐỒỔỖỘƠỚỜỞỠỢƯỨỪỬỮỰ][^:]{3,100}[.?!])$', re.UNICODE)

    for line in text.split("\n"):
        line = line.strip()
        match1 = pattern1.match(line)
        match2 = pattern2.match(line)
        match3 = pattern3.match(line)

        if match1:
            titles.append((match1.group(1), match1.group(3).strip()))
        elif match2:
            titles.append((match2.group(1), match2.group(3).strip()))
        elif match3:
            titles.append((match3.group(1), match3.group(2).strip()))

    return titles

def split_text_into_sections(text, titles):
    """Tách nội dung thành các phần tương ứng với từng tiêu đề."""
    sections = []
    lines = text.split("\n")
    current_title = None
    current_num = None
    current_content = []
    full_titles = [f"{num} {title}".strip() for num, title in titles]
    for line in lines:
        line = line.strip()
        if not line:
            continue
        matched = False
        for i, full_title in enumerate(full_titles):
            if re.match(rf'^{re.escape(full_title)}(\s+|$)', line):
                if current_title is not None:
                    sections.append({
                        "num": current_num,
                        "title": current_title,
                        "content": " ".join(current_content).strip()
                    })
                current_num, current_title = titles[i]
                current_content = []
                matched = True
                break
        if not matched and current_title is not None:
            current_content.append(line)
    if current_title and current_content:
        sections.append({
            "num": current_num,
            "title": current_title,
            "content": " ".join(current_content).strip()
        })
    return sections

def count_words(text):
    """Đếm số từ trong một đoạn văn bản."""
    return len(text.split())

def save_titles_to_excel(sections, output_path):
    """Lưu danh sách tiêu đề và nội dung vào file Excel, kèm số từ, bỏ qua phần có số từ = 0."""
    data = []
    for sec in sections:
        word_count = count_words(sec["content"])
        if word_count < 50:
            continue  # Bỏ qua nếu không có nội dung
        data.append((sec["num"], sec["title"], word_count, sec["content"]))

    if data:
        df = pd.DataFrame(data, columns=["Số thứ tự", "Tiêu đề", "Số từ", "Nội dung"])
        df.to_excel(output_path, index=False)
        print(f"✅ Đã lưu tiêu đề và nội dung vào: {output_path}")
    else:
        print("⚠️ Không có nội dung hợp lệ để lưu.")


def process_pdf(pdf_path):
    """Xử lý từng file PDF và trích xuất dữ liệu."""
    if not os.path.exists(pdf_path):
        print(f"⚠️ File không tồn tại: {pdf_path}")
        return "", [], []

    text = extract_text_from_pdf(pdf_path)
    if not text:
        print(f"⚠️ Không tìm thấy nội dung trong PDF: {pdf_path}")
        return "", [], []

    titles = extract_titles(text)
    sections = split_text_into_sections(text, titles)

    return text, titles, sections

# Danh sách file PDF cần xử lý
pdf_files = [
    r"C:\Tailieuht\TA\Giaotrinh\Đã Làm\lap_trinh_mang.pdf"
]

# Đường dẫn lưu file Excel
output_titles_excel = r"C:\Tailieuht\TA\Giaotrinh\Đã Làm\file.xlsx"

all_sections = []

for pdf_path in pdf_files:
    text, titles, sections = process_pdf(pdf_path)
    if text:
        all_sections.extend(sections)

# Lưu danh sách tiêu đề vào Excel
if all_sections:
    save_titles_to_excel(all_sections, output_titles_excel)
else:
    print("⚠️ Không có phần nội dung nào được trích xuất!")