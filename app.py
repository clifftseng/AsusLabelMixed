import os
import base64
import json
import shutil
import re
import sys
from openai import AzureOpenAI
from dotenv import load_dotenv
import fitz  # PyMuPDF
from openpyxl import load_workbook

# --- Configuration ---
# Load environment variables from .env file
load_dotenv()

# Get Azure OpenAI credentials from environment variables
try:
    AZURE_OPENAI_ENDPOINT = os.environ["AZURE_OPENAI_ENDPOINT"]
    AZURE_OPENAI_API_KEY = os.environ["AZURE_OPENAI_API_KEY"]
    AZURE_OPENAI_DEPLOYMENT_NAME = os.environ["AZURE_OPENAI_DEPLOYMENT"]
    AZURE_OPENAI_API_VERSION = "2024-02-01" # Use a recent API version that supports vision
except KeyError as e:
    print(f"錯誤：請確認您的 .env 檔案中已設定好 {e} 這個環境變數。")
    exit()

# Define directories
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
USER_INPUT_DIR = os.path.join(BASE_DIR, "user_input")
OUTPUT_DIR = os.path.join(BASE_DIR, "output")
PROMPT_DIR = os.path.join(BASE_DIR, "prompt")
EXCEL_OUTPUT_DIR = os.path.join(OUTPUT_DIR, "excel")
SINGLE_TEMPLATE_PATH = os.path.join(BASE_DIR, "single.xlsx")
TOTAL_TEMPLATE_PATH = os.path.join(BASE_DIR, "total.xlsx")


# --- Helper Functions ---
def sanitize_for_excel(text):
    """Removes illegal characters for XML/Excel from a string."""
    if not isinstance(text, str):
        return text
    # XML 1.0 spec forbids characters 0x00-0x08, 0x0B, 0x0C, 0x0E-0x1F
    return re.sub(r'[\x00-\x08\x0B\x0C\x0E-\x1F]', '', text)

def encode_image_to_base64(image_bytes):
    """Encodes image bytes to a base64 string."""
    return base64.b64encode(image_bytes).decode("utf-8")

def pdf_to_base64_images(pdf_path):
    """Converts each page of a PDF to a list of base64 encoded image strings."""
    images = []
    try:
        doc = fitz.open(pdf_path)
        for page_num, page in enumerate(doc):
            pix = page.get_pixmap(dpi=150)
            img_bytes = pix.tobytes("png")
            images.append(encode_image_to_base64(img_bytes))
            print(f"  - 已轉換第 {page_num + 1}/{len(doc)} 頁...")
        doc.close()
    except Exception as e:
        print(f"處理 PDF '{os.path.basename(pdf_path)}' 時發生錯誤: {e}")
        return None
    return images

def read_prompt_file(file_path):
    """Reads content from a prompt file."""
    try:
        with open(file_path, "r", encoding="utf-8") as f:
            return f.read()
    except FileNotFoundError:
        print(f"錯誤：找不到提示檔案 {file_path}")
        return ""

# --- Excel Helper Functions ---
def get_display_value(data_dict):
    """Gets the value to display in Excel, prioritizing raw_value, then derived_value."""
    if not isinstance(data_dict, dict):
        return "無"
    if data_dict.get("raw_value"):
        return data_dict["raw_value"]
    if data_dict.get("derived_value") is not None:
        return f"{data_dict['derived_value']} (推論)"
    return "無"

def format_evidence(evidence_list):
    """Formats the evidence list into a readable string."""
    if not evidence_list:
        return ""
    return "\n".join([
        f"Page {e.get('page', '?')} (loc: {e.get('loc', 'N/A')}): \"{e.get('quote', '')}\""
        for e in evidence_list
    ])

def format_conflicts(conflicts_list):
    """Formats the conflicts list into a readable string."""
    if not conflicts_list:
        return ""
    return json.dumps(conflicts_list, ensure_ascii=False, indent=2)

# --- Main Logic ---
def main():
    """Main function to process PDFs, query Azure OpenAI, and generate Excel reports incrementally."""

    # 1. Initial Setup
    print("--- 清空 output 目錄 ---")
    if os.path.exists(OUTPUT_DIR):
        shutil.rmtree(OUTPUT_DIR)
    os.makedirs(OUTPUT_DIR)
    os.makedirs(EXCEL_OUTPUT_DIR, exist_ok=True)
    print("output 目錄已清空並重建。")

    if not os.path.exists(SINGLE_TEMPLATE_PATH) or not os.path.exists(TOTAL_TEMPLATE_PATH):
        print(f"錯誤: 找不到範本檔案 single.xlsx 或 total.xlsx。請確認檔案位於 {BASE_DIR}")
        return
    
    shutil.copy(TOTAL_TEMPLATE_PATH, os.path.join(EXCEL_OUTPUT_DIR, "total.xlsx"))

    print("\n--- 開始增量處理 PDF 檔案 ---")

    client = AzureOpenAI(
        api_key=AZURE_OPENAI_API_KEY,
        api_version=AZURE_OPENAI_API_VERSION,
        azure_endpoint=AZURE_OPENAI_ENDPOINT
    )
    system_prompt = read_prompt_file(os.path.join(PROMPT_DIR, "prompt_system.txt"))
    user_prompt_template = read_prompt_file(os.path.join(PROMPT_DIR, "prompt_user.txt"))

    if not system_prompt or not user_prompt_template:
        print("錯誤：無法讀取必要的提示檔案，程式終止。")
        return

    pdf_files = [f for f in os.listdir(USER_INPUT_DIR) if f.lower().endswith(".pdf")]
    if not pdf_files:
        print(f"在 {USER_INPUT_DIR} 中找不到任何 PDF 檔案。")
        return

    # 2. Main Incremental Loop
    for filename in pdf_files:
        print(f"\n--- 正在處理檔案: {filename} ---")

        base64_images = pdf_to_base64_images(pdf_path=os.path.join(USER_INPUT_DIR, filename))
        if not base64_images:
            continue

        page_count = len(base64_images)
        current_user_prompt = user_prompt_template.replace("<檔名含副檔名>", filename).replace("<整數>", str(page_count))
        print(f"  - 動態產生使用者提示，檔名: {filename}, 頁數: {page_count}")

        all_json_results = []
        for i in range(0, len(base64_images), 50):
            batch_images = base64_images[i : i + 50]
            print(f"  - 正在處理批次 (頁面 {i+1} 到 {min(i + 50, page_count)})... ")
            
            user_content = [{"type": "text", "text": current_user_prompt}]
            user_content.extend([{"type": "image_url", "image_url": {"url": f"data:image/png;base64,{img}"}} for img in batch_images])

            try:
                print(f"  - 正在向 Azure OpenAI 發送請求 ({len(batch_images)} 張圖片)... ")
                response = client.chat.completions.create(
                    model=AZURE_OPENAI_DEPLOYMENT_NAME,
                    messages=[
                        {"role": "system", "content": system_prompt},
                        {"role": "user", "content": user_content}
                    ],
                    max_tokens=4096, temperature=0.1, top_p=0.95, response_format={"type": "json_object"}
                )
                all_json_results.append(json.loads(response.choices[0].message.content))
                print("  - 成功收到回應。")
            except Exception as e:
                print(f"  - 呼叫 Azure OpenAI API 時發生錯誤: {e}")

        if not all_json_results:
            print(f"檔案 {filename} 沒有產生任何有效的 JSON 結果，跳過後續儲存步驟。")
            continue

        # All subsequent file operations are grouped here for robustness
        try:
            # 1. Prepare and save the JSON data
            merged_json = {k: v for d in all_json_results for k, v in d.items()}
            if 'file' not in merged_json: merged_json['file'] = {}
            merged_json['file']['name'] = filename
            print(f"  - 已使用實際檔案名稱 '{filename}' 覆寫 file.name。")
            
            json_output_filename = os.path.splitext(filename)[0] + ".json"
            with open(os.path.join(OUTPUT_DIR, json_output_filename), "w", encoding="utf-8") as f:
                json.dump(merged_json, f, ensure_ascii=False, indent=4)
            print(f"  - 已儲存 JSON 檔案: {json_output_filename}")

            data = merged_json

            # 2. Generate Single Excel
            single_wb = load_workbook(SINGLE_TEMPLATE_PATH)
            ws = single_wb.active
            ws['B2'] = sanitize_for_excel(data.get('file', {}).get('name', ''))
            ws['B3'] = sanitize_for_excel(data.get('file', {}).get('category', ''))
            
            ws['B4'] = sanitize_for_excel(data.get('model_name', {}).get('value', '無'))
            ws['C4'] = sanitize_for_excel(format_evidence(data.get('model_name', {}).get('evidence', [])))

            fields_map = {
                'nominal_voltage_v': ('B5', 'C5'),
                'typ_batt_capacity_wh': ('B6', 'C6'), 'typ_capacity_mah': ('B7', 'C7'),
                'rated_capacity_mah': ('B8', 'C8'), 'rated_energy_wh': ('B9', 'C9'),
            }
            for key, (val_cell, evi_cell) in fields_map.items():
                field_data = data.get(key, {})
                ws[val_cell] = sanitize_for_excel(get_display_value(field_data))
                ws[evi_cell] = sanitize_for_excel(format_evidence(field_data.get('evidence', [])))
            
            ws['B13'] = sanitize_for_excel(data.get('notes', ''))
            ws['B15'] = sanitize_for_excel(format_conflicts(data.get('conflicts', [])))
            
            excel_filename = os.path.splitext(filename)[0] + ".xlsx"
            single_output_path = os.path.join(EXCEL_OUTPUT_DIR, excel_filename)
            single_wb.save(single_output_path)
            print(f"  - 已儲存單一 Excel 檔案: {excel_filename}")

            # 3. Append to Total Excel and Save
            total_output_path = os.path.join(EXCEL_OUTPUT_DIR, "total.xlsx")
            total_wb = load_workbook(total_output_path)
            total_ws = total_wb.active
            row_data = [
                sanitize_for_excel(data.get('file', {}).get('name', '')), 
                sanitize_for_excel(data.get('file', {}).get('category', '')),
                sanitize_for_excel(data.get('model_name', {}).get('value', '')),
                sanitize_for_excel(get_display_value(data.get('nominal_voltage_v', {}))),
                sanitize_for_excel(get_display_value(data.get('typ_batt_capacity_wh', {}))),
                sanitize_for_excel(get_display_value(data.get('typ_capacity_mah', {}))),
                sanitize_for_excel(get_display_value(data.get('rated_capacity_mah', {}))),
                sanitize_for_excel(get_display_value(data.get('rated_energy_wh', {}))),
                sanitize_for_excel(data.get('notes', '')),
                sanitize_for_excel(format_conflicts(data.get('conflicts', [])))
            ]
            total_ws.append(row_data)
            total_wb.save(total_output_path)
            print(f"  - 已更新並儲存 total.xlsx")

        except Exception as e:
            error_message = f"  - 處理檔案 {filename} 的後續儲存（JSON/Excel）時發生嚴重錯誤: {e}"
            safe_error_message = error_message.encode('utf-8', 'replace').decode(sys.stdout.encoding, 'replace')
            print(safe_error_message)
            continue

    print("\n--- 所有檔案處理完畢 ---")

if __name__ == "__main__":
    main()