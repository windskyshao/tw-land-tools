# -*- coding: utf-8 -*-
"""
高雄市學區查詢模組
從高雄市教育局 PDF 檔案查詢國小、國中學區
"""

import sys
import os

# 🔥 打包環境修正：將 _internal 加入 sys.path
if getattr(sys, 'frozen', False):
    # 如果是從打包的 EXE 執行
    internal_dir = os.path.dirname(os.path.abspath(__file__))
    if internal_dir not in sys.path:
        sys.path.insert(0, internal_dir)
elif '__file__' in globals():
    # 開發環境
    script_dir = os.path.dirname(os.path.abspath(__file__))
    if script_dir not in sys.path:
        sys.path.insert(0, script_dir)


import requests
import re
from io import BytesIO
import pdfplumber
import sys
import io

# 設定 stdout 編碼為 UTF-8
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')


# 高雄市學區劃分 PDF 網址
KAOHSIUNG_ELEMENTARY_PDF = "https://affairs.kh.edu.tw/sites/6236/bulletin_file/29/114%E5%AD%B8%E5%B9%B4%E5%BA%A6%E5%9C%8B%E6%B0%91%E5%B0%8F%E5%AD%B8%E5%AD%B8%E5%8D%80%E5%8A%83%E5%88%86%E4%B8%80%E8%A6%BD%E8%A1%A803.pdf"
KAOHSIUNG_JUNIOR_HIGH_PDF = "https://orgws.kcg.gov.tw/001/KcgOrgUploadFiles/237/relfile/15482/464575/e00246eb-ff50-4aa2-81dc-fcf375bd28d9.pdf"


def download_pdf(url, log_callback=None):
    """
    下載 PDF 檔案

    Args:
        url: PDF 檔案網址
        log_callback: 訊息輸出回調函數

    Returns:
        BytesIO: PDF 檔案內容
    """
    log = log_callback or print

    try:
        log(f"  ⏳ 正在下載 PDF：{url}")
        response = requests.get(url, timeout=30)
        response.raise_for_status()
        log(f"  ✓ PDF 下載完成（{len(response.content)} 位元組）")
        return BytesIO(response.content)
    except Exception as e:
        log(f"  ✗ PDF 下載失敗：{e}")
        return None


def parse_neighbor_list(text):
    """
    解析鄰的列表文字，支援多種格式：
    - 單一鄰：1、3、7
    - 範圍：6-8（表示 6, 7, 8）
    - 混合：1、3、6-8、12-15

    Args:
        text: 包含鄰資訊的文字

    Returns:
        set: 鄰的數字集合
    """
    neighbors = set()

    # 找出所有的鄰表示法，支援：
    # 1. 單一數字+鄰：65鄰
    # 2. 範圍+鄰：65鄰-67鄰 或 65-67鄰
    # 3. 括號內的列表：（1、3、6-8、12-15 鄰）

    # 先處理範圍格式：65鄰-67鄰 或 64-66
    # Pattern 1: 數字-數字（不帶「鄰」），例如：64-66
    range_pattern1 = r'(\d+)[-－](\d+)(?=鄰|、|，|。|\s|）|$)'
    for match in re.finditer(range_pattern1, text):
        start = int(match.group(1))
        end = int(match.group(2))
        neighbors.update(range(start, end + 1))

    # Pattern 2: 數字鄰-數字鄰，例如：65鄰-67鄰
    range_pattern2 = r'(\d+)鄰[-－](\d+)鄰'
    for match in re.finditer(range_pattern2, text):
        start = int(match.group(1))
        end = int(match.group(2))
        neighbors.update(range(start, end + 1))

    # 移除已匹配的範圍，避免重複處理
    text_clean = re.sub(range_pattern1, '', text)
    text_clean = re.sub(range_pattern2, '', text_clean)

    # 處理單一鄰：數字後面可能有「鄰」或標點符號
    single_pattern = r'(\d+)(?=鄰|、|，|。|\s|）|$)'
    for match in re.finditer(single_pattern, text_clean):
        num = int(match.group(1))
        neighbors.add(num)

    return neighbors


def parse_elementary_school_pdf(pdf_file, village, neighbor=None, district=None, log_callback=None):
    """
    解析國小學區 PDF，查詢指定村里（及鄰）對應的國小

    Args:
        pdf_file: PDF 檔案（BytesIO）
        village: 村里名稱（例如：福山里）
        neighbor: 鄰（例如：65 或 065），可選
        district: 行政區（例如：左營區），可選
        log_callback: 訊息輸出回調函數

    Returns:
        list: 國小名稱列表
    """
    log = log_callback or print
    schools = []

    try:
        if neighbor:
            target_neighbor = int(neighbor)
            log(f"  ⏳ 正在解析國小學區 PDF（查詢：{village} 第{target_neighbor}鄰）...")
        else:
            target_neighbor = None
            log(f"  ⏳ 正在解析國小學區 PDF（查詢：{village}）...")

        with pdfplumber.open(pdf_file) as pdf:
            full_text = ""
            for page in pdf.pages:
                page_text = page.extract_text()
                if page_text:
                    full_text += page_text + "\n"

            # 將整個 PDF 文字按行分割
            lines = full_text.split('\n')

            # log(f"  [調試] PDF 總共 {len(lines)} 行")

            # 新策略：PDF是表格格式，學校名稱通常會在單獨一行，然後學區資訊在後續幾行
            # 格式：編號 學校名稱 行政區 學區
            # 例如：「58 福山國小 左營區 福山里（1、3...」

            # 先找所有包含「國小」且沒有「為」「與」的行（學校名稱行）
            school_count = 0
            for i, line in enumerate(lines):
                # 跳過標題行
                if '學區劃分' in line or '學年度' in line:
                    continue

                # 尋找學校名稱：應該包含「國小」但不包含「國民小學」的完整詞
                # 因為表格中學校名稱已經簡化為「福山國小」而非「福山國民小學」
                if '國小' in line and '為' not in line and '與' not in line and '學區' not in line:
                    # 嘗試匹配「編號 + 學校名稱」格式
                    # 例如：「58 福山國小」或「86 文府國小」
                    school_match = re.search(r'^\s*(\d+)\s+([^\s\d]+國小)', line)
                    if not school_match:
                        continue

                    school_number = school_match.group(1)  # 編號
                    school_name = school_match.group(2).strip()  # 學校名稱

                    # 過濾掉無效的學校名稱
                    if len(school_name) < 3 or '學年度' in school_name or '一覽' in school_name:
                        continue

                    school_count += 1
                    # log(f"  [調試] 檢查學校 #{school_count}: {school_name} (第{i}行)")

                    # 找出該學校的所有學區描述
                    # 學區資料可能在學校名稱之前（如文府國小），也可能在之後（如福山國小）
                    # 因此需要找到前一所學校和後一所學校的位置

                    # 找前一所學校的位置
                    section_start = max(0, i - 10)  # 預設往前看10行
                    for j in range(i - 1, max(0, i - 10), -1):
                        if re.match(r'^\s*\d+\s+[^\s\d]+國小', lines[j]):
                            section_start = j + 1  # 從前一所學校的下一行開始
                            break

                    # 找下一所學校的位置
                    section_end = i + 50  # 預設最多看50行
                    for j in range(i + 1, min(i + 50, len(lines))):
                        if re.match(r'^\s*\d+\s+[^\s\d]+國小', lines[j]):
                            section_end = j
                            break

                    # 獲取整個學校的所有資料（包含前面和後面）
                    section_text = '\n'.join(lines[section_start:section_end])

                    # 檢查是否包含目標村里
                    if village not in section_text:
                        continue

                    # 如果指定了行政區，檢查該段落是否包含指定的行政區
                    if district and district not in section_text:
                        continue

                    # log(f"  [調試]   → 學區包含「{village}」，開始檢查鄰")

                    # 如果指定了鄰，進一步檢查
                    if target_neighbor:

                        # 在該段落中找到包含村里的具體描述
                        section_lines = section_text.split('\n')
                        found_match = False

                        for s_line in section_lines:
                            if village in s_line:
                                # 檢查這一行及其附近行是否包含目標鄰
                                idx = section_lines.index(s_line)
                                # log(f"  [調試] 找到村里行{idx}：{s_line[:60]}")

                                # 如果指定了行政區，檢查村里行附近是否有該行政區的標記
                                # 檢查範圍：村里行的前10行到後5行
                                if district:
                                    nearby_lines = section_lines[max(0, idx-10):min(len(section_lines), idx+5)]
                                    nearby_text = '\n'.join(nearby_lines)
                                    # 如果附近沒有目標行政區，跳過這個村里
                                    if district not in nearby_text:
                                        continue

                                context_lines = section_lines[max(0, idx-2):min(len(section_lines), idx+15)]
                                # log(f"  [調試] 提取上下文（{len(context_lines)}行）：")
                                # for ci, cl in enumerate(context_lines):
                                #     log(f"  [調試]   [{ci}] {cl[:60]}")
                                # 移除換行，使學校名稱能跨行匹配
                                context = ''.join(context_lines)
                                # 直接合併所有行（不用換行符），這樣跨行的鄰號才能正確解析
                                neighbors_in_context = parse_neighbor_list(''.join(context_lines))
                                # log(f"  [調試] 上下文中的鄰號（共{len(neighbors_in_context)}個）：{sorted(neighbors_in_context) if len(neighbors_in_context) <=20 else sorted(neighbors_in_context)[:20]}")
                                # log(f"  [調試] 包含65鄰：{target_neighbor in neighbors_in_context}")

                                # log(f"  [調試]   → 找到村里行：{s_line[:50]}")
                                # log(f"  [調試]   → 解析出的鄰: {sorted(list(neighbors_in_context))}")
                                # log(f"  [調試]   → 目標鄰: {target_neighbor}")

                                if target_neighbor in neighbors_in_context:
                                    # 目標鄰在這個上下文的鄰號列表中
                                    # 需要判斷這些鄰號是屬於：
                                    # 1. 當前學校的普通學區
                                    # 2. 自由學區（可選擇多所學校）

                                    # 檢查上下文中是否提到自由學區
                                    # 但要注意：自由學區的鄰號可能和普通學區的鄰號在不同的括號中
                                    # 例如：「福山里（2、4-5...64-66鄰），（20、67鄰，與福山國小自由學區）」
                                    # 這裡64-66是普通學區，20、67才是自由學區

                                    # 尋找包含目標鄰的具體括號段落
                                    # 將文字按括號分段，找所有括號內的內容
                                    bracket_sections = re.findall(r'[（(]([^）)]+)[）)]', '\n'.join(context_lines))

                                    target_is_in_free_zone = False
                                    target_section = None

                                    for section in bracket_sections:
                                        section_neighbors = parse_neighbor_list(section)
                                        if target_neighbor in section_neighbors:
                                            target_section = section
                                            # 檢查這個段落是否包含「自由學區」「共同學區」等字樣
                                            if '自由學區' in section or '共同學區' in section:
                                                target_is_in_free_zone = True
                                            break

                                    if target_is_in_free_zone and target_section:
                                        # 目標鄰在自由學區中，提取所有可選學校
                                        all_schools = re.findall(r'([^、，。：為與\s\n]+?國小)', target_section)
                                        shared_schools = [s for s in all_schools if len(s) >= 3 and '自由' not in s and '學區' not in s]

                                        # 將當前學校和其他共同學區學校都加入
                                        if school_name not in schools:
                                            schools.append(school_name)
                                            log(f"  ✓ 找到國小：{school_name}（自由學區）")

                                        for shared_school in shared_schools:
                                            if shared_school not in schools and shared_school != school_name:
                                                schools.append(shared_school)
                                                log(f"  ✓ 找到國小：{shared_school}（自由學區）")
                                        found_match = True
                                    else:
                                        # 目標鄰在普通學區中，只加入當前學校
                                        if school_name not in schools:
                                            schools.append(school_name)
                                            log(f"  ✓ 找到國小：{school_name}")
                                        found_match = True
                                    break
                                elif not neighbors_in_context:
                                    # 沒有特定鄰的限制，整個村里都適用
                                    if school_name not in schools:
                                        schools.append(school_name)
                                        log(f"  ✓ 找到國小：{school_name}")
                                    found_match = True
                                    break
                    else:
                        # 沒有指定鄰，只要包含村里即可
                        if school_name not in schools:
                            schools.append(school_name)
                            log(f"  ✓ 找到國小：{school_name}")

        if not schools:
            log(f"  ⚠ 未在 PDF 中找到「{village}」第{target_neighbor if target_neighbor else ''}鄰的國小學區")
            # log(f"  [調試] 共檢查了 {school_count} 所學校")

        return schools

    except Exception as e:
        log(f"  ✗ 解析國小學區 PDF 失敗：{e}")
        import traceback
        traceback.print_exc()
        return []


def parse_junior_high_pdf(pdf_file, village, neighbor=None, district=None, log_callback=None):
    """
    解析國中學區 PDF，查詢指定村里（及鄰）對應的國中

    Args:
        pdf_file: PDF 檔案（BytesIO）
        village: 村里名稱（例如：福山里）
        neighbor: 鄰（例如：65 或 065），可選
        district: 行政區（例如：左營區），可選
        log_callback: 訊息輸出回調函數

    Returns:
        list: 國中名稱列表
    """
    log = log_callback or print
    schools = []

    try:
        if neighbor:
            target_neighbor = int(neighbor)
            log(f"  ⏳ 正在解析國中學區 PDF（查詢：{village} 第{target_neighbor}鄰）...")
        else:
            target_neighbor = None
            log(f"  ⏳ 正在解析國中學區 PDF（查詢：{village}）...")

        with pdfplumber.open(pdf_file) as pdf:
            full_text = ""
            for page in pdf.pages:
                page_text = page.extract_text()
                if page_text:
                    full_text += page_text + "\n"

            # 將整個 PDF 文字按行分割
            lines = full_text.split('\n')

            # 新策略：遍歷所有學校，檢查其學區是否包含目標村里+鄰
            for i, line in enumerate(lines):
                # 找到學校名稱行：必須是「編號 + 學校名稱」格式
                # 例如：「11 福山國民中學 左營區」
                if '國民中學' in line and '為' not in line and '與' not in line:
                    # 只匹配有編號開頭的學校行
                    school_match = re.search(r'^\s*\d+\s+([^\s\d、]+?國民中學)', line)
                    if not school_match:
                        continue

                    school_name = school_match.group(1).strip().replace('國民中學', '國中')
                    # log(f"  [調試] 發現學校：{school_name}（原始：{school_match.group(1)}）")

                    # 過濾掉無效的學校名稱
                    if len(school_name) < 3 or '學年度' in school_name or '一覽' in school_name:
                        # log(f"  [調試]   ✗ 過濾掉無效學校")
                        continue

                    # 找出該學校的所有學區描述（與國小相同，需要檢查前後）
                    # 找前一所學校的位置
                    section_start = max(0, i - 10)
                    for j in range(i - 1, max(0, i - 10), -1):
                        if re.match(r'^\s*\d+\s+[^\s\d]+國民中學', lines[j]):
                            section_start = j + 1
                            break

                    # 找下一所學校的位置
                    section_end = i + 50
                    for j in range(i + 1, min(i + 50, len(lines))):
                        if re.match(r'^\s*\d+\s+[^\s\d]+國民中學', lines[j]):
                            section_end = j
                            break

                    section_text = '\n'.join(lines[section_start:section_end])
                    # log(f"  [調試]   檢查段落（行{section_start}-{section_end}）")

                    # 檢查是否包含目標村里
                    if village not in section_text:
                        # log(f"  [調試]   ✗ 未包含村里「{village}」")
                        continue

                    # log(f"  [調試]   ✓ 包含村里「{village}」")

                    # 如果指定了行政區，檢查該段落是否包含指定的行政區
                    if district and district not in section_text:
                        # log(f"  [調試]   ✗ 未包含行政區「{district}」")
                        continue

                    # log(f"  [調試]   ✓ 包含行政區「{district}」")

                    # 如果指定了鄰，進一步檢查
                    if target_neighbor:

                        # 在該段落中找到包含村里的具體描述
                        section_lines = section_text.split('\n')
                        found_match = False

                        for s_line in section_lines:
                            if village in s_line:
                                # 檢查這一行及其附近行是否包含目標鄰
                                idx = section_lines.index(s_line)
                                # log(f"  [調試] 找到村里行{idx}：{s_line[:60]}")

                                # 如果指定了行政區，檢查村里行附近是否有該行政區的標記
                                # 檢查範圍：村里行的前10行到後5行
                                if district:
                                    nearby_lines = section_lines[max(0, idx-10):min(len(section_lines), idx+5)]
                                    nearby_text = '\n'.join(nearby_lines)
                                    # 如果附近沒有目標行政區，跳過這個村里
                                    if district not in nearby_text:
                                        continue

                                context_lines = section_lines[max(0, idx-2):min(len(section_lines), idx+15)]
                                # log(f"  [調試] 提取上下文（{len(context_lines)}行）：")
                                # for ci, cl in enumerate(context_lines):
                                #     log(f"  [調試]   [{ci}] {cl[:60]}")
                                # 移除換行，使學校名稱能跨行匹配
                                context = ''.join(context_lines)
                                # 直接合併所有行（不用換行符），這樣跨行的鄰號才能正確解析
                                neighbors_in_context = parse_neighbor_list(''.join(context_lines))
                                # log(f"  [調試] 上下文中的鄰號（共{len(neighbors_in_context)}個）：{sorted(neighbors_in_context) if len(neighbors_in_context) <=20 else sorted(neighbors_in_context)[:20]}")
                                # log(f"  [調試] 包含65鄰：{target_neighbor in neighbors_in_context}")

                                if target_neighbor in neighbors_in_context:
                                    # 檢查是否為冒號形式的自由學區（例如：2鄰、4鄰...65鄰-67鄰：為文府國中與福山國中自由學區）
                                    # 需要找到包含village名稱的那一段，然後從該段往後找冒號模式
                                    context_text = ''.join(context_lines)

                                    # 找到包含village的那一行開始的文本
                                    village_pattern = f'{village}[^{village}]{{0,200}}?[:：]\\s*為([^、]+國中)[與、]([^、\\s]+國中)(?:自由學區|共同學區)'
                                    colon_free_zone_match = re.search(village_pattern, context_text)

                                    if colon_free_zone_match:
                                        # 這是冒號形式的自由學區，提取從村里名稱到冒號之間的文字
                                        full_match = colon_free_zone_match.group(0)
                                        # 從村里名稱開始到冒號之間的部分
                                        village_to_colon = full_match.split('：')[0].split(':')[0]
                                        zone_neighbors = parse_neighbor_list(village_to_colon)

                                        # log(f"  [調試] 檢查冒號形式自由學區：{full_match[:100]}")
                                        # log(f"  [調試] 鄰號區段：{village_to_colon[:80]}")
                                        # log(f"  [調試] 解析到的鄰號：{sorted(zone_neighbors) if len(zone_neighbors)<=20 else sorted(zone_neighbors)[:20]}")

                                        if target_neighbor in zone_neighbors:
                                            # 提取所有學校名稱
                                            all_schools = re.findall(r'([^、，。：為與\s\n]+?國中)', full_match)
                                            shared_schools = [s for s in all_schools if len(s) >= 3 and '自由' not in s and '學區' not in s]

                                            # log(f"  [調試] 找到冒號形式自由學區（包含65鄰）")
                                            # log(f"  [調試] 提取到的學校：{shared_schools}")

                                            if school_name not in schools:
                                                schools.append(school_name)
                                                log(f"  ✓ 找到國中：{school_name}（自由學區）")

                                            for shared_school in shared_schools:
                                                if shared_school not in schools and shared_school != school_name:
                                                    schools.append(shared_school)
                                                    log(f"  ✓ 找到國中：{shared_school}（自由學區）")
                                            found_match = True
                                            break

                                    # 使用與國小相同的括號分段邏輯
                                    # 尋找包含目標鄰的具體括號段落，判斷是普通學區還是自由學區
                                    bracket_sections = re.findall(r'[（(]([^）)]+)[）)]', '\n'.join(context_lines))

                                    # log(f"  [調試] 學校：{school_name}，找到{len(bracket_sections)}個括號段落")

                                    target_is_in_free_zone = False
                                    target_section = None

                                    if bracket_sections:
                                        # 有括號段落，逐一檢查
                                        for idx, section in enumerate(bracket_sections):
                                            # log(f"  [調試] 學校「{school_name}」段落{idx+1}: {section[:60]}...")
                                            section_neighbors = parse_neighbor_list(section)
                                            # log(f"  [調試]   鄰號: {sorted(section_neighbors)[:15]}")
                                            # log(f"  [調試]   包含65: {target_neighbor in section_neighbors}")
                                            if target_neighbor in section_neighbors:
                                                target_section = section
                                                # log(f"  [調試]   有自由學區: {'自由學區' in section or '共同學區' in section}")
                                                if '自由學區' in section or '共同學區' in section:
                                                    target_is_in_free_zone = True
                                                break

                                        if target_is_in_free_zone and target_section:
                                            # 目標鄰在自由學區中，提取所有可選學校
                                            # log(f"  [調試] 處理自由學區，段落內容: {target_section}")
                                            all_schools = re.findall(r'([^、，。：為與\s\n]+?國中)', target_section)
                                            # log(f"  [調試] 提取到的學校: {all_schools}")
                                            shared_schools = [s for s in all_schools if len(s) >= 3 and '自由' not in s and '學區' not in s]
                                            # log(f"  [調試] 過濾後的學校: {shared_schools}")

                                            if school_name not in schools:
                                                schools.append(school_name)
                                                log(f"  ✓ 找到國中：{school_name}（自由學區）")

                                            for shared_school in shared_schools:
                                                if shared_school not in schools and shared_school != school_name:
                                                    schools.append(shared_school)
                                                    log(f"  ✓ 找到國中：{shared_school}（自由學區）")
                                            found_match = True
                                        elif target_section:
                                            # 找到包含目標鄰的括號段落，但不是自由學區
                                            if school_name not in schools:
                                                schools.append(school_name)
                                                log(f"  ✓ 找到國中：{school_name}")
                                            found_match = True
                                    else:
                                        # 沒有括號段落，但目標鄰在上下文中 = 普通學區
                                        if school_name not in schools:
                                            schools.append(school_name)
                                            log(f"  ✓ 找到國中：{school_name}")
                                        found_match = True
                                    break
                                elif not neighbors_in_context:
                                    # 沒有特定鄰的限制，整個村里都適用
                                    if school_name not in schools:
                                        schools.append(school_name)
                                        log(f"  ✓ 找到國中：{school_name}")
                                    found_match = True
                                    break
                    else:
                        # 沒有指定鄰，只要包含村里即可
                        if school_name not in schools:
                            schools.append(school_name)
                            log(f"  ✓ 找到國中：{school_name}")

        if not schools:
            log(f"  ⚠ 未在 PDF 中找到「{village}」的國中學區")

        return schools

    except Exception as e:
        log(f"  ✗ 解析國中學區 PDF 失敗：{e}")
        import traceback
        traceback.print_exc()
        return []


def format_school_list(schools, school_type):
    """
    格式化學校列表輸出

    Args:
        schools: 學校名稱列表（例如：['文府國中', '福山國中']）
        school_type: 學校類型（'國小' 或 '國中'）

    Returns:
        str: 格式化後的字串（例如：'文府、福山國中'）
    """
    if not schools:
        return ""

    # 移除每個學校名稱中的類型後綴
    cleaned_schools = []
    for school in schools:
        # 移除「國小」或「國中」後綴
        cleaned = school.replace('國民小學', '').replace('國民中學', '').replace('國小', '').replace('國中', '')
        cleaned_schools.append(cleaned)

    # 用「、」連接學校名稱，最後加上類型
    if len(cleaned_schools) == 1:
        return f"{cleaned_schools[0]}{school_type}"
    else:
        return '、'.join(cleaned_schools) + school_type


def query_kaohsiung_schools(village, neighbor=None, district=None, log_callback=None):
    """
    查詢高雄市指定村里（及鄰）的國小、國中學區

    Args:
        village: 村里名稱（例如：福山里）
        neighbor: 鄰（例如：65 或 065），可選
        district: 行政區（例如：左營區），可選，用於過濾學校
        log_callback: 訊息輸出回調函數

    Returns:
        dict: {'國小': [...], '國中': [...]}
    """
    log = log_callback or print

    result = {
        '國小': [],
        '國中': []
    }

    log(f"\n[步驟 7] 查詢高雄市學區...")
    if neighbor:
        log(f"  村里：{village}  鄰：第{neighbor}鄰")
    else:
        log(f"  村里：{village}")
    if district:
        log(f"  行政區：{district}")

    # 下載並解析國小學區 PDF
    elementary_pdf = download_pdf(KAOHSIUNG_ELEMENTARY_PDF, log_callback)
    if elementary_pdf:
        result['國小'] = parse_elementary_school_pdf(elementary_pdf, village, neighbor, district, log_callback)

    # 下載並解析國中學區 PDF
    junior_high_pdf = download_pdf(KAOHSIUNG_JUNIOR_HIGH_PDF, log_callback)
    if junior_high_pdf:
        result['國中'] = parse_junior_high_pdf(junior_high_pdf, village, neighbor, district, log_callback)

    # 輸出結果
    if result['國小'] or result['國中']:
        log(f"\n✓ 學區查詢完成：")
        if result['國小']:
            formatted = format_school_list(result['國小'], '國小')
            log(f"  國小：{formatted}")
        if result['國中']:
            formatted = format_school_list(result['國中'], '國中')
            log(f"  國中：{formatted}")
    else:
        log(f"\n⚠ 未查詢到學區資料")

    return result


if __name__ == "__main__":
    # 測試：左營區福山里第65鄰
    result = query_kaohsiung_schools("福山里", neighbor="65", district="左營區")
    print(f"\n結果：{result}")

    # 驗證結果
    expected_elementary = ['文府國小']
    expected_junior = ['文府國中', '福山國中']

    if result['國小'] == expected_elementary and set(result['國中']) == set(expected_junior):
        print("\n✓✓✓ 測試通過！")
    else:
        print("\n✗✗✗ 測試失敗！")
        print(f"預期國小：{expected_elementary}，實際：{result['國小']}")
        print(f"預期國中：{expected_junior}，實際：{result['國中']}")
