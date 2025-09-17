import os
import base64
import json
from openai import AzureOpenAI
from dotenv import load_dotenv
import fitz  # PyMuPDF

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
    print(f"錯誤：請確認您的 .env 檔案中已設定好 {e} 這個環境變數。 সন")
    exit()

# Define directories
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
USER_INPUT_DIR = os.path.join(BASE_DIR, "user_input")
OUTPUT_DIR = os.path.join(BASE_DIR, "output")
PROMPT_DIR = os.path.join(BASE_DIR, "prompt")

# --- Helper Functions ---
def encode_image_to_base64(image_bytes):
    """Encodes image bytes to a base64 string."""
    return base64.b64encode(image_bytes).decode("utf-8")

def pdf_to_base64_images(pdf_path):
    """Converts each page of a PDF to a list of base64 encoded image strings."""
    images = []
    try:
        doc = fitz.open(pdf_path)
        for page_num, page in enumerate(doc):
            # Render page to a pixmap (an image)
            pix = page.get_pixmap(dpi=150) # Use a reasonable DPI
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

# --- Main Logic ---
def main():
    """Main function to process PDFs and query Azure OpenAI."""
    print("--- 開始處理 PDF 檔案 ---")

    # Ensure output directory exists
    os.makedirs(OUTPUT_DIR, exist_ok=True)

    # Initialize Azure OpenAI client
    client = AzureOpenAI(
        api_key=AZURE_OPENAI_API_KEY,
        api_version=AZURE_OPENAI_API_VERSION,
        azure_endpoint=AZURE_OPENAI_ENDPOINT
    )

    # Read prompts
    system_prompt = read_prompt_file(os.path.join(PROMPT_DIR, "prompt_system.txt"))
    user_prompt_template = read_prompt_file(os.path.join(PROMPT_DIR, "prompt_user.txt"))

    if not system_prompt or not user_prompt_template:
        print("錯誤：無法讀取必要的提示檔案，程式終止。")
        return

    # Process each PDF in the user_input directory
    for filename in os.listdir(USER_INPUT_DIR):
        if filename.lower().endswith(".pdf"):
            pdf_path = os.path.join(USER_INPUT_DIR, filename)
            print(f"\n正在處理檔案: {filename}")

            # Convert PDF pages to images
            base64_images = pdf_to_base64_images(pdf_path)
            if not base64_images:
                continue

            all_json_results = [] # To store JSON results from all batches
            image_batch_size = 50 # Max images per request

            for i in range(0, len(base64_images), image_batch_size):
                batch_num = (i // image_batch_size) + 1
                batch_images = base64_images[i : i + image_batch_size]
                print(f"  - 正在處理第 {batch_num} 批次 (頁面 {i+1} 到 {min(i + image_batch_size, len(base64_images))})... ")

                # Construct the message payload for the Vision API
                user_content = [{"type": "text", "text": user_prompt_template.encode('utf-8').decode('utf-8')}]
                for img in batch_images:
                    user_content.append({
                        "type": "image_url",
                        "image_url": {"url": f"data:image/png;base64,{img}"}
                    })

                messages = [
                    {"role": "system", "content": system_prompt.encode('utf-8').decode('utf-8')},
                    {"role": "user", "content": user_content}
                ]

                # Call Azure OpenAI API
                try:
                    print(f"  - 本次請求包含 {len(user_content)} 個內容元素 (1 個文字提示 + {len(user_content) - 1} 張圖片)。")
                    print("  - 正在向 Azure OpenAI 發送請求...")
                    response = client.chat.completions.create(
                        model=AZURE_OPENAI_DEPLOYMENT_NAME,
                        messages=messages,
                        max_tokens=4096, # Adjust as needed
                        temperature=0.1,
                        top_p=0.95,
                        response_format={"type": "json_object"} # Request JSON output
                    )

                    # Extract JSON content
                    result_content = response.choices[0].message.content
                    print("  - 成功收到回應。 সন")
                    all_json_results.append(json.loads(result_content))

                except Exception as e:
                    print(f"  - 呼叫 Azure OpenAI API 時發生錯誤: {e}")
                    # Continue to next batch or next file if an error occurs

            # --- Merge JSON results ---
            if all_json_results:
                merged_json = {}
                for res_dict in all_json_results:
                    # Simple merge: update dictionary. Last value for a key wins.
                    merged_json.update(res_dict)

                # Save the merged JSON response
                output_filename = os.path.splitext(filename)[0] + ".json"
                output_path = os.path.join(OUTPUT_DIR, output_filename)

                try:
                    with open(output_path, "w", encoding="utf-8") as f:
                        json.dump(merged_json, f, ensure_ascii=False, indent=4)
                    print(f"所有批次結果已合併並儲存至: {output_path}")
                except Exception as e:
                    print(f"儲存合併後的 JSON 時發生錯誤: {e}")
            else:
                print(f"檔案 {filename} 沒有產生任何有效的 JSON 結果。")

    print("\n--- 所有檔案處理完畢 ---")

if __name__ == "__main__":
    main()