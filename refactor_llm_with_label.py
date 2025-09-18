import re
import os
import datetime

APP_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "app.py")

NEW_FUNC = r'''
def method_llm_with_label():
    """
    Processes PDF files based on matching JSON formats, generates images,
    then immediately sends processed images to Azure OpenAI for labeling and saves JSON per file.
    """
    print("--- 開始根據 format JSON 處理圖片並呼叫 Azure OpenAI（每檔案即時呼叫與儲存） ---")

    # Ensure output directory exists
    if not os.path.exists(OUTPUT_DIR):
        os.makedirs(OUTPUT_DIR)
        print(f"已建立 output 目錄: {OUTPUT_DIR}")

    # 1) 建立 format name -> json 路徑（不分大小寫）
    format_dir = os.path.join(BASE_DIR, "format")
    try:
        format_files = [f for f in os.listdir(format_dir) if f.lower().endswith('.json')]
        format_map = {os.path.splitext(f)[0].lower(): os.path.join(format_dir, f) for f in format_files}
        print(f"成功載入 {len(format_map)} 個 format JSON 檔案。")
    except FileNotFoundError:
        print(f"錯誤: 找不到 format 目錄: {format_dir}")
        return

    # 2) 初始化 AOAI client 與 system prompt（只做一次）
    client = AzureOpenAI(
        api_key=AZURE_OPENAI_API_KEY,
        api_version=AZURE_OPENAI_API_VERSION,
        azure_endpoint=AZURE_OPENAI_ENDPOINT
    )
    system_prompt_aoai_path = os.path.join(PROMPT_DIR, "prompt_system_using_label.txt")
    system_prompt_aoai = read_prompt_file(system_prompt_aoai_path)

    if not system_prompt_aoai:
        print(f"錯誤：找不到或無法讀取 Azure OpenAI 的系統提示檔案 {system_prompt_aoai_path}，程式終止。")
        return

    # 3) 掃描 user_input 內的 PDF
    pdf_files = [f for f in os.listdir(USER_INPUT_DIR) if f.lower().endswith(".pdf")]
    if not pdf_files:
        print(f"在 {USER_INPUT_DIR} 中找不到任何 PDF 檔案。")
        return

    print(f"找到 {len(pdf_files)} 個 PDF 檔案，開始進行比對與處理...")

    for pdf_filename in pdf_files:
        pdf_base_name = os.path.splitext(pdf_filename)[0]
        matched_format_key = None

        # 以 regex 優先：完整詞邊界比對
        for format_key in format_map.keys():
            if re.search(r'\b' + re.escape(format_key) + r'\b', pdf_filename, re.IGNORECASE):
                matched_format_key = format_key
                break

        # 若 regex 沒命中，再退而求其次用 in
        if not matched_format_key:
            for format_key in format_map.keys():
                if format_key in pdf_filename.lower():
                    matched_format_key = format_key
                    break

        if not matched_format_key:
            print(f"\n- 檔案 '{pdf_filename}' 未匹配到任何格式，已跳過。")
            continue

        print(f"\n- 處理檔案 '{pdf_filename}' (匹配到格式: '{matched_format_key}')")
        json_path = format_map[matched_format_key]
        pdf_path = os.path.join(USER_INPUT_DIR, pdf_filename)

        pdf_output_subdir = os.path.join(OUTPUT_DIR, pdf_base_name)
        os.makedirs(pdf_output_subdir, exist_ok=True)
        print(f"  - 已為檔案 '{pdf_filename}' 建立輸出子目錄: {pdf_output_subdir}")

        doc = None
        try:
            # 讀入相對應的 format 設定
            with open(json_path, 'r', encoding='utf-8') as f:
                config = json.load(f)

            doc = fitz.open(pdf_path)
            if not doc.page_count > 0:
                print("  - 警告: PDF 為空，無法處理。")
                continue

            # (A) 依 width/height 由第 1 頁產生縮圖
            max_width = config.get('width')
            max_height = config.get('height')

            if max_width and max_height:
                first_page = doc[0]
                pix_first_page = first_page.get_pixmap(dpi=200)
                original_image_p1 = Image.frombytes("RGB", [pix_first_page.width, pix_first_page.height], pix_first_page.samples)

                resized_image = original_image_p1.copy()
                resized_image.thumbnail((max_width, max_height), Image.Resampling.LANCZOS)

                resized_filename = f"{pdf_base_name}_resized.png"
                resized_path = os.path.join(pdf_output_subdir, resized_filename)
                resized_image.save(resized_path)
                print(f"  - 已儲存縮放後的圖片 (第一頁): {resized_filename}")
            else:
                print("  - 警告: JSON 中缺少 'width' 或 'height' 設定，跳過第一頁縮放。")

            # (B) 依 hints 切圖（逐頁）
            if 'hints' in config and isinstance(config['hints'], list):
                for hint in config['hints']:
                    page_num = hint.get('page')
                    field_name = hint.get('field')
                    bbox = hint.get('bbox')

                    if not (page_num and field_name and bbox and isinstance(page_num, int) and page_num > 0):
                        print("  - 警告: 'hints' 中的項目格式不正確或缺少 'page'/'field'/'bbox'。跳過此 hint。")
                        continue

                    if page_num > doc.page_count:
                        print(f"  - 警告: hint 指定的頁面 {page_num} 超出 PDF 總頁數 {doc.page_count}。跳過此 hint。")
                        continue

                    target_page = doc[page_num - 1]
                    pix_target_page = target_page.get_pixmap(dpi=200)
                    image_to_crop = Image.frombytes("RGB", [pix_target_page.width, pix_target_page.height], pix_target_page.samples)

                    if not (isinstance(bbox, list) and len(bbox) == 4):
                        print(f"  - 警告: field '{field_name}' 的 bbox 格式不正確，預期為 [x, y, w, h] 陣列。跳過切割。")
                        continue

                    x, y, w, h = bbox[0], bbox[1], bbox[2], bbox[3]
                    crop_box = (x, y, x + w, y + h)

                    if crop_box[0] < 0 or crop_box[1] < 0 or crop_box[2] > image_to_crop.width or crop_box[3] > image_to_crop.height:
                        print(f"  - 警告: field '{field_name}' 的 bbox ({x},{y},{w},{h}) 超出頁面 {page_num} 的圖片範圍 ({image_to_crop.width}x{image_to_crop.height})，跳過切割。")
                        continue

                    cropped_image = image_to_crop.crop(crop_box)
                    cropped_filename = f"{field_name}.png"
                    cropped_path = os.path.join(pdf_output_subdir, cropped_filename)
                    cropped_image.save(cropped_path)
                    print(f"  - 已切割並儲存 (頁面 {page_num}, 欄位 '{field_name}'): {cropped_filename}")
            else:
                print("  - 警告: JSON 中沒有 'hints' 列表或其為空，跳過圖片切割。")

        except Exception as e:
            print(f"  - 處理檔案 '{pdf_filename}' 時發生錯誤: {e}")
            # 若圖片處理就失敗，直接處理下一份 PDF
            if doc:
                try:
                    doc.close()
                except:
                    pass
            continue
        finally:
            if doc:
                try:
                    doc.close()
                except:
                    pass

        # === 重點：在每一份 PDF 的圖片流程結束後，立刻呼叫 AOAI 並儲存 JSON ===
        # 收集該 PDF 子目錄下的所有 PNG，轉成 base64
        image_files = [f for f in os.listdir(pdf_output_subdir) if os.path.isfile(os.path.join(pdf_output_subdir, f)) and f.lower().endswith(".png")]
        if not image_files:
            print(f"  - 在 '{pdf_output_subdir}' 中找不到任何圖片檔案，跳過 Azure OpenAI 請求。")
            continue

        base64_images_for_aoai = []
        for img_file in image_files:
            img_path = os.path.join(pdf_output_subdir, img_file)
            base64_img = image_file_to_base64(img_path)
            if base64_img:
                base64_images_for_aoai.append(base64_img)

        if not base64_images_for_aoai:
            print(f"  - 無法編碼 '{pdf_output_subdir}' 中的任何圖片，跳過 Azure OpenAI 請求。")
            continue

        user_content_aoai = [
            {"type": "text", "text": "請根據提供的圖片，提取所有相關資訊，並以 JSON 格式回應。"}
        ]
        user_content_aoai.extend([
            {"type": "image_url", "image_url": {"url": f"data:image/png;base64,{img}"}}
            for img in base64_images_for_aoai
        ])

        try:
            print(f"  - 正在向 Azure OpenAI 發送請求 ({len(base64_images_for_aoai)} 張圖片)... ")
            response = client.chat.completions.create(
                model=AZURE_OPENAI_DEPLOYMENT_NAME,
                messages=[
                    {"role": "system", "content": system_prompt_aoai},
                    {"role": "user", "content": user_content_aoai}
                ],
                max_tokens=4096, temperature=0.1, top_p=0.95, response_format={"type": "json_object"}
            )
            aoai_json_response = json.loads(response.choices[0].message.content)
            print("  - 成功收到 Azure OpenAI 回應。" )

            # 以 <PDF 基名>_with_label.json 存在該 PDF 的輸出子目錄
            output_json_filename = f"{pdf_base_name}_with_label.json"
            output_json_path = os.path.join(pdf_output_subdir, output_json_filename)
            with open(output_json_path, "w", encoding="utf-8") as f:
                json.dump(aoai_json_response, f, ensure_ascii=False, indent=4)
            print(f"  - 已儲存 Azure OpenAI 回應: {output_json_filename}")

        except Exception as e:
            print(f"  - 呼叫 Azure OpenAI API 或處理回應時發生錯誤: {e}")

    print("\n--- 所有檔案均已完成：圖片處理 -> 即時 AOAI 呼叫 -> 即時 JSON 儲存 ---")
'''

def main():
    if not os.path.exists(APP_PATH):
        raise FileNotFoundError(f"找不到 app.py：{APP_PATH}")

    with open(APP_PATH, "r", encoding="utf-8") as f:
        src = f.read()

    # 以「def method_llm_with_label( ... ) 到 if __name__ == "__main__": 之前」為界做整段替換
    pattern = re.compile(
        r"(def\s+method_llm_with_label\s*\([\s\S]*?)\n(?=if\s+__name__\s*==\s*['"]__main__['"]\s*:)",
        re.DOTALL
    )

    if not pattern.search(src):
        raise RuntimeError("找不到 method_llm_with_label() 函式區塊，無法替換。")

    new_src = pattern.sub(NEW_FUNC.rstrip("\n") + "\n", src)

    # 建立備份
    ts = datetime.datetime.now().strftime("%Y%m%d-%H%M%S")
    backup_path = APP_PATH + f".bak-{ts}"
    with open(backup_path, "w", encoding="utf-8") as f:
        f.write(src)

    with open(APP_PATH, "w", encoding="utf-8") as f:
        f.write(new_src)

    print(f"已完成替換並寫回 app.py。備份檔：{backup_path}")

if __name__ == "__main__":
    main()
