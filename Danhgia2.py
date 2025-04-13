from transformers import T5ForConditionalGeneration, T5Tokenizer
import pandas as pd
from sklearn.metrics.pairwise import cosine_similarity
import torch

# Tải mô hình T5 từ Hugging Face
model_name = "t5-small"
tokenizer = T5Tokenizer.from_pretrained(model_name)
model = T5ForConditionalGeneration.from_pretrained(model_name)

# Hàm tính điểm độ tương đồng (semantic similarity)
def get_similarity(text1, text2):
    inputs = tokenizer(text1, return_tensors="pt", padding=True, truncation=True)
    targets = tokenizer(text2, return_tensors="pt", padding=True, truncation=True)
    
    with torch.no_grad():
        # Tính ma trận similarity giữa hai văn bản
        input_embeddings = model.encoder(inputs.input_ids)[0].mean(dim=1)
        target_embeddings = model.encoder(targets.input_ids)[0].mean(dim=1)
    
    sim_score = cosine_similarity(input_embeddings.numpy(), target_embeddings.numpy())
    return sim_score[0][0]

# Hàm đánh giá tính trung thực, mạch lạc, liên quan
def evaluate_summary(article, summary):
    # Tính độ tương đồng (Faithfulness)
    faith_score = get_similarity(article, summary)
    # Chuyển điểm độ tương đồng (0-1) thành thang điểm 5
    faith_score = round(faith_score * 5, 2)

    if faith_score >= 4.5:
        faith_comment = "Hoàn toàn trung thực, không có sự khác biệt."
    elif 3.5 <= faith_score < 4.5:
        faith_comment = "Bản tóm tắt khá trung thực, có thể bỏ sót một số chi tiết."
    elif 2.5 <= faith_score < 3.5:
        faith_comment = "Bản tóm tắt thiếu trung thực, nhiều thông tin bị sai lệch."
    else:
        faith_comment = "Thông tin sai lệch nghiêm trọng."

    # Đánh giá tính mạch lạc (Coherence)
    coh_score = 5 if len(summary.split()) > 50 else 4
    coh_comment = "Mạch lạc, nhưng có thể cải thiện sự liên kết giữa các câu." if coh_score == 4 else "Hoàn toàn mạch lạc."

    # Đánh giá tính liên quan (Relevance)
    if len(summary.split()) > len(article.split()) / 2:
        rel_score = 5
        rel_comment = "Bản tóm tắt giữ lại các ý chính."
    else:
        rel_score = 4
        rel_comment = "Có một số chi tiết không quan trọng."

    # Tính điểm trung bình
    avg_score = round((faith_score + coh_score + rel_score) / 3, 2)

    return {
        "Điểm tính trung thực": faith_score,
        "Nhận xét tính trung thực": faith_comment,
        "Điểm tính mạch lạc": coh_score,
        "Nhận xét mạch lạc": coh_comment,
        "Điểm tính liên quan": rel_score,
        "Nhận xét tính liên quan": rel_comment,
        "Điểm trung bình cộng": avg_score,
        "Nhận xét chung": f"{faith_comment} {coh_comment} {rel_comment}"
    }

# Hàm xử lý file Excel
def process_files(input_file, output_file):
    df = pd.read_excel(input_file)

    # Kiểm tra nếu file không có các cột cần thiết
    required_columns = ["Đoạn văn", "Tóm tắt", "Tiêu đề đoạn văn", "Tên giáo trình", "Tác giả"]
    for column in required_columns:
        if column not in df.columns:
            print(f"❌ Không tìm thấy cột '{column}' trong file.")
            return

    for index, row in df.iterrows():
        result = evaluate_summary(row["Đoạn văn"], row["Tóm tắt"])
        df.at[index, "Điểm tính trung thực"] = result["Điểm tính trung thực"]
        df.at[index, "Nhận xét tính trung thực"] = result["Nhận xét tính trung thực"]
        df.at[index, "Điểm tính mạch lạc"] = result["Điểm tính mạch lạc"]
        df.at[index, "Nhận xét mạch lạc"] = result["Nhận xét mạch lạc"]
        df.at[index, "Điểm tính liên quan"] = result["Điểm tính liên quan"]
        df.at[index, "Nhận xét tính liên quan"] = result["Nhận xét tính liên quan"]
        df.at[index, "Điểm trung bình cộng"] = result["Điểm trung bình cộng"]
        df.at[index, "Nhận xét chung"] = result["Nhận xét chung"]

    # Lưu kết quả vào file Excel
    df.to_excel(output_file, index=False, engine="openpyxl")
    print(f"✅ Đã lưu kết quả đánh giá vào: {output_file}")

# Đường dẫn file input và output
input_file = r"C:\Tailieuht\TA\ex\Exam.xlsx"  # Thay đổi đường dẫn file của bạn
output_file = r"C:\Tailieuht\TA\ex\Exam.xlsx" # Kết quả sẽ được lưu ở đây

process_files(input_file, output_file)
