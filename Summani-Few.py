import pandas as pd
import google.generativeai as genai
import os
import time

def generate_summary(article, summarizer_model, summarization_prompt, retries=3):
    user_prompt = f"""
    {summarization_prompt}

    Văn bản gốc:
    {article}

    Hãy viết một bản tóm tắt ngắn gọn, rõ ràng và đầy đủ thông tin chính.
    """

    attempt = 0
    while attempt < retries:
        try:
            response = summarizer_model.generate_content(
                user_prompt,
                generation_config=genai.GenerationConfig(
                    temperature=0.3,
                    top_p=0.9,
                    top_k=50
                )
            )
            return response.text.strip() if response and hasattr(response, "text") else "No response"
        except Exception as e:
            attempt += 1
            print(f"❌ Lỗi khi tóm tắt (lần thử {attempt}/{retries}): {e}")
            time.sleep(30)  # chờ 30s trước khi thử lại
    return "LỖI khi tóm tắt"

def process_files(input_file, output_file, gemini_api_key, summarizer_id, summarization_prompt_file):
    df = pd.read_excel(input_file)

    print("✅ Các cột trong DataFrame:", df.columns)

    genai.configure(api_key=gemini_api_key)

    with open(summarization_prompt_file, "r", encoding="utf-8") as f:
        summarization_prompt = f.read()

    summarizer_model = genai.GenerativeModel(
        model_name=summarizer_id,
        system_instruction=summarization_prompt
    )

    summaries = []
    for idx, article in enumerate(df['Văn bản gốc']):
        print(f"🔄 Đang xử lý dòng {idx + 1}/{len(df)}")
        summary = generate_summary(article, summarizer_model, summarization_prompt)
        summaries.append(summary)

    df['Văn bản tóm tắt'] = summaries

    try:
        df.to_excel(output_file, index=False)
        print(f"✅ Đã lưu kết quả vào {output_file}")
    except Exception as e:
        print(f"⚠️ Lỗi khi lưu file chính: {e}")
        backup_file = output_file.replace(".xlsx", "_backup.xlsx")
        try:
            df.to_excel(backup_file, index=False)
            print(f"✅ Đã lưu file backup vào {backup_file}")
        except Exception as backup_error:
            print(f"❌ Không thể lưu backup: {backup_error}")

# Cấu hình đường dẫn và API key
input_file = r"C:\Tailieuht\TA\ex\test-fewshot.xlsx"
output_file = r"C:\Tailieuht\TA\ex\test-fewshot.xlsx"
gemini_api_key = "AIzaSyBnSYyjnK8gRQ_j41qA7uAvk-jqHSiqmOE"
summarizer_id = "gemini-1.5-flash-latest"  # Sử dụng Flash-Lite
summarization_prompt_file = r"C:\Tailieuht\TA\ex\sumani-fewshot.txt"

process_files(input_file, output_file, gemini_api_key, summarizer_id, summarization_prompt_file)
