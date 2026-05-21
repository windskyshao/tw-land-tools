# -*- coding: utf-8 -*-
"""
建立學區資料結構
從 PDF 解析學區資料，建立結構化的 JSON 檔案
"""
import sys
import io

# 🔥 在 PyInstaller 環境中，sys.stdout 可能是 None，需要檢查
if sys.stdout is not None and hasattr(sys.stdout, 'buffer'):
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

import pdfplumber
import re
import json
import os

# 🔥 引入 base_dir_helper 以取得正確的輸出路徑
try:
    from base_dir_helper import get_internal_dir
except ImportError:
    # 如果無法引入，使用當前目錄
    def get_internal_dir():
        return os.path.dirname(os.path.abspath(__file__))


def parse_neighbor_list(text):
    """
    解析鄰號列表
    例如: "1鄰、3鄰、7-10鄰、65-67鄰"
    返回: [1, 3, 7, 8, 9, 10, 65, 66, 67]

    支援的格式：
    - 單一鄰號: "1鄰"、"3鄰"
    - 範圍格式1: "7-10鄰" (結尾有鄰)
    - 範圍格式2: "2鄰-5鄰" (兩端都有鄰)
    - 範圍格式3: "34-40鄰" (結尾有鄰，但開頭沒有)
    """
    neighbors = set()

    # 移除換行和多餘空白
    text = text.replace('\n', '').replace(' ', '')

    # 🔥 改進的範圍解析
    # 格式1: "2鄰-5鄰" -> 匹配 "數字鄰-數字鄰"
    range_pattern1 = r'(\d+)鄰[-–—](\d+)鄰'
    for match in re.finditer(range_pattern1, text):
        start = int(match.group(1))
        end = int(match.group(2))
        neighbors.update(range(start, end + 1))

    # 格式2: "7-10鄰" 或 "34-40鄰" -> 匹配 "數字-數字鄰"（開頭不能是「鄰」）
    range_pattern2 = r'(?<!鄰)(\d+)[-–—](\d+)鄰'
    for match in re.finditer(range_pattern2, text):
        start = int(match.group(1))
        end = int(match.group(2))
        neighbors.update(range(start, end + 1))

    # 🔥 格式3: "1-3、" 或 "1-3," -> 匹配「數字-數字」後面跟著分隔符（不是鄰字）
    # 這種格式常見於 "1-3、25-44鄰" 這類組合表示法
    range_pattern3 = r'(\d+)[-–—](\d+)(?=[、,，])'
    for match in re.finditer(range_pattern3, text):
        start = int(match.group(1))
        end = int(match.group(2))
        neighbors.update(range(start, end + 1))

    # 處理單一鄰號: "1鄰"、"3鄰"
    # 排除已經在範圍中處理過的（前面或後面有連字符）
    single_pattern = r'(?<![0-9])(\d+)鄰(?![-–—])'
    for match in re.finditer(single_pattern, text):
        # 檢查這個數字前面是否有連字符（表示是範圍的結尾）
        pos = match.start()
        if pos > 0 and text[pos-1] in '-–—':
            continue  # 跳過，這是範圍的結尾
        neighbor_num = int(match.group(1))
        neighbors.add(neighbor_num)

    return sorted(list(neighbors))


def extract_free_zone_info(area_data):
    """
    提取自由學區資訊
    返回: [(鄰號列表, 學校列表), ...]

    🔥 改進：只提取「自由學區」前面緊鄰的鄰號，避免提取到「除XX鄰為YY學區外」的鄰號
    """
    free_zones = []

    # 🔥 先移除換行符，避免學校名稱被截斷（如「壽\n山國中」→「壽山國中」）
    area_data = area_data.replace('\n', '')

    # 🔥 改進的策略：
    # 1. 先用「；」或「、」分割句子
    # 2. 只處理包含「自由學區」或「共同學區」的片段
    # 3. 排除包含「除」「外」的排除語句

    # 按分號分割，處理每個片段
    segments = re.split(r'[；;]', area_data)

    for segment in segments:
        # 只處理包含「自由學區」的片段
        if '自由學區' not in segment and '共同學區' not in segment:
            continue

        # 🔥 如果這個片段包含「除...外」的排除語句，需要特殊處理
        # 例如：「除14鄰、15鄰為XX學區外，另20鄰-24鄰為YY自由學區」
        if '除' in segment and '外' in segment:
            # 找到「另」或「其中」後面的部分，那才是自由學區的鄰號
            # 或者找到最接近「自由學區」的鄰號
            match = re.search(r'(?:另|其中)?(\d+鄰[-–—\d鄰、，,\s]+)(?:為|：)[^自由學區]*自由學區', segment)
            if match:
                neighbor_text = match.group(1)
                neighbors = parse_neighbor_list(neighbor_text)
            else:
                # 如果沒有明確的「另」，嘗試提取最接近「自由學區」的鄰號
                # 從「自由學區」往前找鄰號
                free_zone_pos = segment.find('自由學區')
                if free_zone_pos > 0:
                    # 往前找到「為」或「：」
                    before_text = segment[:free_zone_pos]
                    # 找最後一個鄰號範圍
                    neighbor_matches = list(re.finditer(r'(\d+鄰[-–—\d鄰、，,\s]*)', before_text))
                    if neighbor_matches:
                        # 取最後一個匹配
                        last_match = neighbor_matches[-1]
                        neighbor_text = last_match.group(1)
                        neighbors = parse_neighbor_list(neighbor_text)
                    else:
                        neighbors = []
                else:
                    neighbors = []
        else:
            # 沒有排除語句，正常解析
            neighbors = parse_neighbor_list(segment)

        # 提取學校名稱（國中或國小）
        school_pattern = r'([^、，。：為與\s\n]+?國(?:中|小)(?:部)?)'
        schools = re.findall(school_pattern, segment)
        schools = [s for s in schools if len(s) >= 3 and '自由' not in s and '學區' not in s]

        # 🔥 如果有學校但沒有鄰號，表示「全里」都是自由學區，用 [-1] 標記
        if schools:
            if not neighbors:
                neighbors = [-1]  # 全里是自由學區
            free_zones.append({
                'neighbors': neighbors,
                'schools': schools,
                'text': segment.strip()
            })

    return free_zones


def parse_notes_section(text, school_data):
    """
    解析備註區的資訊
    例如: "瑞南里（1-3、25-44鄰）"
    會更新 school_data 中對應村里的鄰號資訊
    """
    if not text or '備註' not in text:
        return

    # 提取備註內容（在「備註：」或「備註:」之後）
    notes_pattern = r'備註\s*[：:]\s*(.+)'
    notes_match = re.search(notes_pattern, text, re.DOTALL)

    if not notes_match:
        return

    notes_content = notes_match.group(1)
    print(f"    📝 發現備註區: {notes_content[:100]}...")

    # 解析備註中的村里和鄰號
    # 格式: "村里名（鄰號範圍）"
    village_with_neighbors_pattern = r'([^、，。：（）\s]+里)\s*[（(]\s*([^）)]+)\s*[）)]'
    matches = re.findall(village_with_neighbors_pattern, notes_content)

    for village_name, neighbor_text in matches:
        neighbors = parse_neighbor_list(neighbor_text)
        if neighbors:
            print(f"      ✓ 備註中的村里: {village_name} -> 鄰號: {neighbors}")
            # 在所有學校和行政區中尋找這個村里，並更新鄰號
            for school_name in school_data:
                for district in school_data[school_name]:
                    if village_name in school_data[school_name][district]:
                        existing = set(school_data[school_name][district][village_name]['regular'])
                        existing.update(neighbors)
                        school_data[school_name][district][village_name]['regular'] = sorted(list(existing))
                        print(f"        → 已更新 {school_name}/{district}/{village_name}")


def normalize_school_name(raw_name, school_type='國中'):
    """
    標準化學校名稱
    處理各種格式：國民中學、高中國中部、國民小學等
    """
    if not raw_name:
        return None

    # 移除換行和空白
    name = raw_name.replace('\n', '').replace(' ', '').strip()

    if school_type == '國中':
        # 檢查是否為國中
        if '國民中學' in name:
            # 標準國民中學 -> 國中
            return name.replace('國民中學', '國中')
        elif '高中' in name and '國中部' in name:
            # 高中國中部（如：文山高中國中部、福誠高中國中部）
            # 保留完整名稱但標準化
            return name.replace('國中部', '國中部')
        elif '高級中學' in name and '國中部' in name:
            # 高級中學國中部
            return name.replace('高級中學', '高中')
        elif '國中' in name:
            # 已經是簡化格式
            return name
    elif school_type == '國小':
        if '國民小學' in name:
            return name.replace('國民小學', '國小')
        elif '國中小' in name:
            # 🔥 國中小學校（如：翠屏國中小、巴楠花部落國中小）
            # 如果沒有 (國小部) 後綴，加上它
            if '(國小部)' not in name:
                return name + '(國小部)'
            return name
        elif '國小' in name:
            return name

    return None


def clean_village_name(village_name):
    """
    清理村里名稱，移除不正確的前綴和後綴
    """
    if not village_name:
        return None

    name = village_name.strip()

    # 🔥 移除開頭的標點符號（如 "。南成里" -> "南成里"、".中正里" -> "中正里"）
    name = re.sub(r'^[.。,，、;；:：\s]+', '', name)

    # 移除括號及其內容如果在開頭（如「(全里」）
    if name.startswith('(') or name.startswith('（'):
        return None

    # 如果名稱不以「里」結尾，返回 None
    if not name.endswith('里'):
        return None

    # 🔥 只有當第一個字是行政區前綴，且後面是有效的村里名稱時才移除
    # 例如：「區文英里」-> 「文英里」
    # 但：「鎮北里」-> 保持「鎮北里」（「鎮北」是村里名稱的一部分）
    if name[0] == '區' and len(name) >= 3:
        # 「區」後面通常不是村里名稱的一部分，移除它
        name = name[1:]

    # 🔥 過濾掉不是真正村里名稱的情況
    # - 「全里」不是村里名稱
    # - 「鄰XX里」不是村里名稱
    # - 「以上X里」、「以下X里」不是村里名稱（如「以上5里」、「以上2里」）
    # - 包含「鄰」字的不是村里名稱（如「19鄰盛昌里」、「16鄰福安里」）
    # - 包含多個「里」的情況（如「大華里鳥松里」）需要特殊處理
    # - 🔥 包含備註文字的不是村里名稱（如「起增加湖東里」）
    if name == '全里':
        return None
    if name.startswith('鄰'):
        return None
    # 🔥 過濾「以上X里」、「以下X里」格式
    if re.match(r'^以[上下]\d+里$', name):
        return None
    # 🔥 過濾「數字+鄰+村里名」格式（如「19鄰盛昌里」、「16鄰福安里」）
    # 這是鄰號被誤連接到村里名稱的情況
    # 但要保留正常的村里名稱（如「福安里」本身不應該被過濾）
    if re.match(r'^\d+鄰', name):
        return None
    # 🔥 過濾包含備註關鍵字的名稱（如「起增加湖東里」、「三民區鼎泰里」）
    note_keywords = ['增加', '調整', '劃歸', '變更', '原為', '改為', '自由學區']
    for kw in note_keywords:
        if kw in name:
            return None
    # 🔥 過濾包含「行政區+里」格式的誤解析（如「三民區鼎泰里」）
    # 正常村里名稱不應該包含「區」字
    if '區' in name and name.index('區') > 0:
        return None

    # 檢查是否有多個「里」字（可能是兩個村里連在一起）
    li_count = name.count('里')
    if li_count > 1:
        # 嘗試分割，取最後一個村里
        # 例如：「大華里鳥松里」-> 不處理，由呼叫者處理
        return None

    # 移除名稱中的特殊字符
    name = re.sub(r'[【】\[\]（()）]', '', name)

    # 確保名稱長度合理（2-5個字）
    if len(name) < 2 or len(name) > 6:
        return None

    return name


def parse_text_based_school_data(page_text, school_data, school_type, current_school_name, current_district, last_village_name=None):
    """
    🔥 純文字解析：當表格無法識別時使用
    適用於 115 年國小學區等格式變更的 PDF

    返回: (更新後的 current_school_name, 更新後的 current_district, 更新後的 last_village_name)
    """
    if not page_text:
        return current_school_name, current_district, last_village_name

    lines = page_text.split('\n')

    # 🔥 115 年國小 PDF 格式：
    # "1 旗津國小 旗津區 踐里、旗下里、北汕里..."
    # "1 4 鎮昌國小 前鎮區 平昌里、平等里..."（數字之間有空格）
    # "2 8 七賢國小" （沒有行政區，在下一行）

    # 正則表達式匹配學校行（有行政區）
    # 格式: 編號(可能有空格，如 "1 4" 或 "1 07") + 學校名(含國小/國中小) + 行政區 + 學區內容
    # 🔥 支援三位數編號，如 "1 07"=107, "1 60"=160
    # 🔥 支援國中小(國小部)、分校等特殊名稱
    # 🔥 學校名稱應為 2-10 個中文字（巴楠花部落國中小 有 8 個字）
    school_name_pattern = r'(?:[\u4e00-\u9fff]{2,10}(?:國小|國中小)(?:\(國小部\))?|[\u4e00-\u9fff]{2,10}國小\S*分校)'

    # 🔥 高雄市行政區列表（用於驗證）
    valid_districts = (
        '鹽埕區|鼓山區|左營區|楠梓區|三民區|新興區|前金區|苓雅區|前鎮區|旗津區|小港區|'
        '鳳山區|林園區|大寮區|大樹區|大社區|仁武區|鳥松區|岡山區|橋頭區|燕巢區|田寮區|'
        '阿蓮區|路竹區|湖內區|茄萣區|永安區|彌陀區|梓官區|旗山區|美濃區|六龜區|甲仙區|'
        '杉林區|內門區|茂林區|桃源區|那瑪夏區'
    )

    school_pattern_full = re.compile(
        r'^(\d+(?:\s+\d+){0,2})\s+'  # 編號（可能像 "1 4" 或 "1 07" 這樣有空格）
        rf'({school_name_pattern})\s+'  # 學校名稱
        rf'({valid_districts})\s*'  # 🔥 只匹配真正的行政區名稱
        r'(.*)$'  # 學區內容（可選）
    )

    # 🔥 匹配只有編號和學校名稱的行（行政區在下一行）
    # 支援三位數編號
    school_pattern_short = re.compile(
        r'^(\d+(?:\s+\d+){0,2})\s+'  # 編號
        rf'({school_name_pattern})$'  # 學校名稱（結尾）
    )

    # 🔥 匹配編號+學校名稱+非行政區的內容（如 "1 0 瑞豐國小 里、瑞興里"）
    # 這種情況是跨行的學區資料，不是新學校的主要條目
    school_pattern_with_area = re.compile(
        r'^(\d+(?:\s+\d+){0,2})\s+'  # 編號
        rf'({school_name_pattern})\s+'  # 學校名稱
        r'(.+)$'  # 後面的內容（可能是學區資料或任何內容）
    )

    # 🔥 匹配特殊格式：學校名稱 + 行政區（沒有編號）
    # 如 "岡山國小 岡山區" 或 "內門國小 內門區"
    school_pattern_nonum = re.compile(
        rf'^({school_name_pattern})\s+'  # 學校名稱
        rf'({valid_districts})\s*'  # 🔥 只匹配真正的行政區名稱
        r'(.*)$'  # 學區內容（可選）
    )

    # 🔥 匹配只有學校名稱的行（如 "岡山國小" 單獨一行）
    # 注意：學校名稱應該是 2-10 個中文字 + "國小" 或 "國中小"
    school_pattern_only = re.compile(
        r'^([\u4e00-\u9fff]{2,10}(?:國小|國中小)(?:\(國小部\))?)$'  # 只有學校名稱
    )

    # 匹配只有行政區和學區內容的行（延續上一個學校）
    # 🔥 使用前面定義的 valid_districts
    district_pattern = re.compile(
        rf'^({valid_districts})\s+'  # 只匹配真正的行政區名稱
        r'(.+)$'  # 學區內容
    )

    for line in lines:
        line = line.strip()
        if not line:
            continue

        # 跳過標題行
        if '編 號' in line or '學校名稱' in line or '學 區' in line:
            continue

        # 🔥 先嘗試匹配完整的學校行（有行政區）
        match = school_pattern_full.match(line)
        if match:
            school_num, school_name, district, area_data = match.groups()

            # 標準化學校名稱
            school_name_clean = school_name.replace(' ', '').strip()
            normalized = normalize_school_name(school_name_clean, school_type)

            # 🔥 過濾無效的學校名稱（包含「里」「鄰」等地名元素的不是學校名）
            if normalized and not any(kw in normalized for kw in ['里', '鄰']):
                current_school_name = normalized
                current_district = district.strip()

                if current_school_name not in school_data:
                    school_data[current_school_name] = {}
                    print(f"    發現學校(文字解析): {current_school_name}")

                # 處理學區資料
                if current_district not in school_data[current_school_name]:
                    school_data[current_school_name][current_district] = {}

                # 解析村里資料
                if area_data:
                    last_village_name = process_area_data_text(area_data, school_data, current_school_name, current_district, school_type, last_village_name)
            continue

        # 🔥 嘗試匹配編號+學校名稱+內容（但內容不是行政區）
        # 這種情況可能是：
        # 1. 學區資料的延續（如 "1 0 瑞豐國小 里、瑞興里"）
        # 2. 有學校但後面是非標準格式的學區資料
        match_with_area = school_pattern_with_area.match(line)
        if match_with_area:
            school_num, school_name, remaining = match_with_area.groups()

            # 🔥 檢查 remaining 是否以有效行政區開頭（而不是任何以「區」結尾的字串）
            remaining_stripped = remaining.strip()
            is_district = bool(re.match(rf'^({valid_districts})', remaining_stripped))

            if not is_district:
                # 這是一個新學校，但後面跟著學區資料而不是行政區
                school_name_clean = school_name.replace(' ', '').strip()
                normalized = normalize_school_name(school_name_clean, school_type)

                if normalized and not any(kw in normalized for kw in ['里', '鄰']):
                    current_school_name = normalized
                    # 🔥 嘗試從 remaining 中提取行政區
                    # 格式可能是 "(6-7、12鄰，與十全國小自由學區)、興德里、"
                    # 或 "內坑里(與永芳國小自由學區)"
                    district_in_remaining = re.search(rf'({valid_districts})', remaining)
                    if district_in_remaining:
                        current_district = district_in_remaining.group(1)
                    # 🔥 如果 remaining 中沒有行政區，使用之前記住的行政區（從前一行獲取的）
                    # 不要將 current_district 設為 None

                    if current_school_name not in school_data:
                        school_data[current_school_name] = {}
                        print(f"    發現學校(編號+名稱+資料): {current_school_name}")

                    # 🔥 如果有行政區，處理學區資料
                    if current_district:
                        if current_district not in school_data[current_school_name]:
                            school_data[current_school_name][current_district] = {}
                        last_village_name = process_area_data_text(remaining, school_data, current_school_name, current_district, school_type, last_village_name)
                continue

        # 🔥 嘗試匹配只有編號和學校名稱的行（行政區在下一行）
        match_short = school_pattern_short.match(line)
        if match_short:
            school_num, school_name = match_short.groups()

            school_name_clean = school_name.replace(' ', '').strip()
            normalized = normalize_school_name(school_name_clean, school_type)

            if normalized:
                current_school_name = normalized
                # 行政區會在下一行，先清空
                current_district = None

                if current_school_name not in school_data:
                    school_data[current_school_name] = {}
                    print(f"    發現學校(文字解析): {current_school_name}")
            continue

        # 🔥 嘗試匹配沒有編號的學校行（如 "岡山國小 岡山區" 或 "內門國小 內門區"）
        match_nonum = school_pattern_nonum.match(line)
        if match_nonum:
            school_name, district, area_data = match_nonum.groups()

            school_name_clean = school_name.replace(' ', '').strip()
            normalized = normalize_school_name(school_name_clean, school_type)

            # 🔥 過濾無效的學校名稱（包含「里」「鄰」等地名元素的不是學校名）
            if normalized and not any(kw in normalized for kw in ['里', '鄰']):
                current_school_name = normalized
                current_district = district.strip()

                if current_school_name not in school_data:
                    school_data[current_school_name] = {}
                    print(f"    發現學校(無編號): {current_school_name}")

                if current_district not in school_data[current_school_name]:
                    school_data[current_school_name][current_district] = {}

                if area_data:
                    last_village_name = process_area_data_text(area_data, school_data, current_school_name, current_district, school_type, last_village_name)
            continue

        # 🔥 嘗試匹配只有學校名稱的行（如 "岡山國小" 單獨一行）
        match_only = school_pattern_only.match(line)
        if match_only:
            school_name = match_only.group(1)

            school_name_clean = school_name.replace(' ', '').strip()
            normalized = normalize_school_name(school_name_clean, school_type)

            # 🔥 過濾無效的學校名稱（包含「里」「鄰」「區」等地名元素的不是學校名）
            if normalized and not any(kw in normalized for kw in ['里', '鄰', '區']):
                current_school_name = normalized
                current_district = None  # 行政區會在下一行

                if current_school_name not in school_data:
                    school_data[current_school_name] = {}
                    print(f"    發現學校(僅名稱): {current_school_name}")
            continue

        # 🔥 嘗試匹配被截斷的「國中小」學校名稱（如 "巴楠花部落國" 應該是 "巴楠花部落國中小"）
        # 格式: XX國 或 XXX國 結尾（被截斷的國中小）
        truncated_match = re.match(r'^([\u4e00-\u9fff]{3,8}國)$', line.strip())
        if truncated_match and school_type == '國小':
            truncated_name = truncated_match.group(1)
            # 假設這是被截斷的「國中小」名稱
            full_name = truncated_name + '中小(國小部)'

            if full_name not in school_data:
                school_data[full_name] = {}
                current_school_name = full_name
                current_district = None
                print(f"    發現學校(截斷名): {full_name}")
            continue

        # 嘗試匹配只有行政區的行
        district_match = district_pattern.match(line)
        if district_match:
            district, area_data = district_match.groups()
            pending_district = district.strip()  # 🔥 記錄待處理的行政區

            # 如果當前有學校，這可能是延續上一學校
            if current_school_name:
                # 🔥 判斷是延續上一學校還是為下一學校準備
                # 如果 area_data 包含里名，則是延續
                # 如果 area_data 主要是 "區）" 等殘留，則可能是新行政區
                if '里' in area_data or '鄰' in area_data:
                    current_district = pending_district
                    if current_district not in school_data[current_school_name]:
                        school_data[current_school_name][current_district] = {}
                    last_village_name = process_area_data_text(area_data, school_data, current_school_name, current_district, school_type, last_village_name)
                else:
                    # 這個行政區可能是為下一個學校準備的
                    current_district = pending_district
            else:
                current_district = pending_district
            continue

        # 如果這行看起來像是學區資料的延續（包含「里」或「鄰」）
        if current_school_name and current_district and ('里' in line or '鄰' in line):
            # 檢查是否不是備註行
            if not any(kw in line for kw in ['備註', '學年度', '調整為', '自由學區']):
                last_village_name = process_area_data_text(line, school_data, current_school_name, current_district, school_type, last_village_name)

    return current_school_name, current_district, last_village_name


def process_area_data_text(area_data, school_data, school_name, district, school_type, last_village_name=None):
    """
    處理純文字解析的學區資料

    返回: 最後處理的村里名稱（用於跨頁續接）
    """
    if not area_data or not school_name or not district:
        return last_village_name

    # 清理資料
    area_data = area_data.replace('\n', '').strip()

    # 提取村里名稱和鄰號
    # 格式如：「中華里、永安里、振興里(1-5鄰)、復興里」

    # 先分割自由學區資訊
    free_zones = extract_free_zone_info(area_data)

    # 🔥 先檢查是否是「孤立鄰號」（跨頁續接情況）
    # 格式如: "（1-3、25-44鄰）。" 或 "(1-3、25-44鄰)"
    # 這種情況沒有「里」字，但有鄰號括號
    orphan_neighbor_match = re.match(r'^[\s（(]+([^)）里]+鄰[^)）]*)[\)）]', area_data)
    if orphan_neighbor_match and last_village_name and '里' not in area_data:
        neighbor_info = orphan_neighbor_match.group(1)
        neighbors = parse_neighbor_list(neighbor_info)
        if neighbors and last_village_name in school_data[school_name][district]:
            existing = school_data[school_name][district][last_village_name].get('regular', [])
            school_data[school_name][district][last_village_name]['regular'] = sorted(list(set(existing + neighbors)))
            print(f"      🔗 跨頁續接: {last_village_name} 新增鄰號 {neighbors}")
        return last_village_name

    # 使用正則找出所有「XX里」或「XX里(鄰號)」
    village_pattern = re.compile(r'([^\d、，,\s（()）\[\]【】]+里)(?:\s*[\(（]([^)）]+)[\)）])?')

    current_last_village = last_village_name

    for match in village_pattern.finditer(area_data):
        village_name = match.group(1)
        neighbor_info = match.group(2)

        # 標準化村里名稱
        normalized_village = clean_village_name(village_name)
        if not normalized_village:
            continue

        # 記住最後處理的村里名稱
        current_last_village = normalized_village

        # 確保村里存在於資料結構中
        if normalized_village not in school_data[school_name][district]:
            school_data[school_name][district][normalized_village] = {
                'regular': [],
                'free_zones': []
            }

        # 解析鄰號
        if neighbor_info:
            neighbors = parse_neighbor_list(neighbor_info + '鄰')
            if neighbors:
                existing = school_data[school_name][district][normalized_village].get('regular', [])
                school_data[school_name][district][normalized_village]['regular'] = sorted(list(set(existing + neighbors)))
        else:
            # 沒有指定鄰號，表示全里
            if not school_data[school_name][district][normalized_village].get('regular'):
                school_data[school_name][district][normalized_village]['regular'] = []

    # 處理自由學區
    for free_zone in free_zones:
        neighbors = free_zone.get('neighbors', [])
        shared_schools = free_zone.get('schools', [])
        # 找出這些鄰號屬於哪個村里
        for village_name in school_data[school_name][district]:
            village_data = school_data[school_name][district][village_name]
            if neighbors:
                village_data['free_zones'].append({
                    'neighbors': neighbors,
                    'shared_with': shared_schools
                })

    return current_last_village


def is_school_name(text, school_type='國中'):
    """
    判斷文字是否為學校名稱
    """
    if not text:
        return False

    text = text.replace('\n', '').replace(' ', '').strip()

    if school_type == '國中':
        # 檢查各種國中格式
        if '國民中學' in text:
            return True
        if '高中' in text and '國中部' in text:
            return True
        if '高級中學' in text and '國中部' in text:
            return True
        # 特殊學校：國立高雄餐旅大學附屬高級中學國中部、國立高雄師大附中國中部等
        if '附中國中部' in text or '附屬' in text and '國中部' in text:
            return True
    elif school_type == '國小':
        if '國民小學' in text or '國小' in text:
            return True

    return False


def build_school_district_data(pdf_path, school_type='國中'):
    """
    從 PDF 建立學區資料結構

    資料結構:
    {
        "學校名稱": {
            "行政區": {
                "村里名": {
                    "regular": [鄰號列表],  # 普通學區
                    "free_zones": [         # 自由學區
                        {
                            "neighbors": [鄰號],
                            "shared_with": [其他學校]
                        }
                    ]
                }
            }
        }
    }
    """
    school_data = {}

    # 🔥 修正跨頁問題：將 current_school_name 和 current_district 移到外層
    # 這樣可以在處理跨頁表格時保持狀態
    current_school_name = None
    current_district = None
    previous_area_data = None  # 🔥 保存上一行的村里資料，用於檢測跨頁拆分的村里名稱
    pending_school_name_parts = []  # 🔥 用於累積跨行的學校名稱
    last_village_name = None  # 🔥 記住最後處理的村里名稱，用於跨頁續接

    print(f"正在解析 {school_type} PDF: {pdf_path}")

    with pdfplumber.open(pdf_path) as pdf:
        for page_num, page in enumerate(pdf.pages, 1):
            print(f"  處理第 {page_num} 頁...")

            # 提取整頁文字，用於解析備註區
            page_text = page.extract_text()
            if page_text and '備註' in page_text:
                parse_notes_section(page_text, school_data)

            tables = page.extract_tables()

            if not tables:
                # 🔥 表格無法識別時，使用純文字解析（適用於 115 年國小學區等格式）
                if school_type == '國小' and page_text:
                    print(f"    (使用純文字解析模式)")
                    current_school_name, current_district, last_village_name = parse_text_based_school_data(
                        page_text, school_data, school_type,
                        current_school_name, current_district, last_village_name
                    )
                continue

            for table in tables:
                current_school = None
                # 🔥 不再重置 current_school_name 和 current_district
                # current_school_name = None
                # current_district = None

                for row_idx, row in enumerate(table):
                    if not row:
                        continue

                    # 處理不同的表格格式
                    if len(row) == 4:
                        # 國中格式
                        school_num, school_name, row_district, area_data = row
                    elif len(row) >= 10:
                        # 國小格式
                        school_num = row[1] if len(row) > 1 else None
                        school_name = row[3] if len(row) > 3 else None
                        row_district = row[6] if len(row) > 6 else None
                        area_data = row[9] if len(row) > 9 else None
                    else:
                        continue

                    # 🔥 改進的學校名稱識別邏輯
                    # 處理跨行學校名稱和各種格式（國民中學、高中國中部等）
                    if school_name and school_name.strip():
                        school_name_cleaned = school_name.replace('\n', '').replace(' ', '').strip()

                        # 檢查是否為完整的學校名稱
                        if is_school_name(school_name_cleaned, school_type):
                            # 如果有累積的部分，先合併
                            if pending_school_name_parts:
                                full_name = ''.join(pending_school_name_parts) + school_name_cleaned
                                pending_school_name_parts = []
                                normalized = normalize_school_name(full_name, school_type)
                            else:
                                normalized = normalize_school_name(school_name_cleaned, school_type)

                            if normalized:
                                current_school_name = normalized
                                if current_school_name not in school_data:
                                    school_data[current_school_name] = {}
                                    print(f"    發現學校: {current_school_name}")
                        else:
                            # 可能是學校名稱的一部分（跨行），累積起來
                            # 但只有當它看起來像是學校名稱的開頭時才累積
                            if any(keyword in school_name_cleaned for keyword in ['國立', '市立', '縣立', '高中', '高級']):
                                pending_school_name_parts.append(school_name_cleaned)
                            elif pending_school_name_parts:
                                # 繼續累積
                                pending_school_name_parts.append(school_name_cleaned)
                                # 檢查累積後是否為完整學校名稱
                                combined = ''.join(pending_school_name_parts)
                                if is_school_name(combined, school_type):
                                    normalized = normalize_school_name(combined, school_type)
                                    if normalized:
                                        current_school_name = normalized
                                        if current_school_name not in school_data:
                                            school_data[current_school_name] = {}
                                            print(f"    發現學校（跨行）: {current_school_name}")
                                        pending_school_name_parts = []

                    # 更新當前行政區
                    if row_district and row_district.strip():
                        current_district = row_district.strip()

                    # 🔥 跨頁處理：只要有 area_data 且有 current_school_name 和 current_district 就處理
                    # 這樣可以處理跨頁表格中只有「區」欄位有資料的情況
                    if not area_data or not area_data.strip():
                        continue

                    if not current_school_name or not current_district:
                        # 如果沒有學校或行政區，跳過（但這不應該發生，因為我們保持了狀態）
                        continue

                    # 確保行政區存在
                    if current_district not in school_data[current_school_name]:
                        school_data[current_school_name][current_district] = {}

                    # 🔥 解決跨頁村里名稱拆分問題
                    # 如果上一行的 area_data 以不完整的村里名稱結尾，且當前行開頭是「xx里」
                    # 則合併它們（例如：上一行「瑞」+ 當前行「和里」= 「瑞和里」）
                    if previous_area_data and previous_area_data.strip():
                        previous_cleaned = previous_area_data.replace('\n', '').strip()
                        last_char = previous_cleaned[-1] if previous_cleaned else ''

                        # 🔥 判斷是否需要合併的條件：
                        # 1. 上一行最後一個字是中文
                        # 2. 上一行不是以「里」、「鄰」、「區」、標點符號結尾
                        # 3. 當前行以村里名稱開頭（如「XX里」）
                        # 4. 🔥 重要：上一行不能包含「里」字（表示已經是完整的村里列表）
                        #    這樣可以避免「維生里...全市韓僑學生」的「生」被誤合併
                        should_merge = (
                            last_char
                            and '\u4e00' <= last_char <= '\u9fff'
                            and last_char not in '里鄰區）)】'
                            and not last_char.isdigit()
                            and '里' not in previous_cleaned  # 🔥 如果上一行有「里」字，表示是完整資料
                        )

                        if should_merge:
                            area_data = last_char + area_data
                            print(f"    🔗 合併跨頁村里：{last_char} + {area_data[:20]}...")

                    # 保存當前的 area_data 供下一次使用
                    previous_area_data = area_data

                    # 🔥 處理儲存格內的多行資料
                    # 1. 先處理學校名稱中的換行符（如「壽\n山國中」→「壽山國中」）
                    #    這種情況是學校名稱被 PDF 表格分行截斷
                    # 2. 再按換行符分割，分別處理每一行，然後用頓號合併
                    #    這樣可以避免「大華里\n鳥松里14鄰」變成「大華里鳥松里14鄰」

                    # 🔥 先修復被換行符截斷的名稱
                    # PDF 表格中常見的問題：名稱被分行截斷
                    # 例如：「壽\n山國中」、「麗興\n里」、「登\n山里」
                    area_data_fixed = area_data

                    # 修復村里名稱中的換行符
                    # 情況1：「XX\n里」→「XX里」（如「麗興\n里」→「麗興里」）
                    area_data_fixed = re.sub(r'(\w+)\n(里)', r'\1\2', area_data_fixed)
                    # 情況2：「X\nYY里」→「XYY里」（如「登\n山里」→「登山里」）
                    # 🔥 只有當換行符前面是「中文字」（不是里、鄰、數字、括號）時才合併
                    # 避免「興隆里\n合和里」、「19鄰\n祿興里」、「（中山高速公路以東）\n南成里」被誤合併
                    # 使用 lookbehind 確保前面是中文字但不是「里」、「鄰」或括號
                    area_data_fixed = re.sub(r'(?<=[^\d里鄰）\)\]】])\n(\w+里)', r'\1', area_data_fixed)

                    # 修復學校名稱中的換行符
                    # 「X\nY國中」格式（X 是學校名稱的一部分）
                    area_data_fixed = re.sub(r'(\w)\n(\w+國中(?:部)?)', r'\1\2', area_data_fixed)
                    # 「X\nY高中」格式
                    area_data_fixed = re.sub(r'(\w)\n(\w+高中)', r'\1\2', area_data_fixed)

                    area_lines = area_data_fixed.split('\n')
                    # 合併時用頓號分隔，讓每個村里可以被正確識別
                    area_data_cleaned = '、'.join(line.strip() for line in area_lines if line.strip())

                    # 解析村里和鄰號
                    # 策略：先找出所有村里名稱及其在文字中的位置，然後提取每個村里到下一個村里之間的文字
                    # 🔥 簡化村里名稱模式：匹配「XX里」格式
                    village_pattern = r'([^、，。：\s（()【\[\]】]+里)'

                    # 找出所有村里及其位置
                    village_matches = []
                    for match in re.finditer(village_pattern, area_data_cleaned):
                        raw_village_name = match.group(1)
                        # 排除包含這些關鍵字的假村里名稱
                        if '自由' in raw_village_name or '學區' in raw_village_name or '共同' in raw_village_name:
                            continue

                        # 🔥 清理村里名稱
                        village_name = clean_village_name(raw_village_name)
                        if village_name:
                            village_matches.append({
                                'name': village_name,
                                'start': match.start(),
                                'end': match.end()
                            })

                    # 🔥 檢測「以上X里」模式：如果整段文字包含「以上X里」，
                    # 則自由學區資訊應用到前面所有列出的村里
                    shared_free_zones = None
                    if re.search(r'以上\d+里', area_data_cleaned):
                        # 從整段文字中提取自由學區資訊
                        shared_free_zones = extract_free_zone_info(area_data_cleaned)

                    # 對每個村里，提取從它開始到下一個村里開始之間的文字
                    for i, village_info in enumerate(village_matches):
                        village = village_info['name']

                        # 確保村里存在
                        if village not in school_data[current_school_name][current_district]:
                            school_data[current_school_name][current_district][village] = {
                                'regular': [],
                                'free_zones': []
                            }

                        # 提取這個村里的文字段落（從當前村里到下一個村里，或到結尾）
                        start_pos = village_info['start']
                        if i + 1 < len(village_matches):
                            # 還有下一個村里，提取到下一個村里開始的位置
                            end_pos = village_matches[i + 1]['start']
                        else:
                            # 這是最後一個村里，提取到結尾
                            end_pos = len(area_data_cleaned)

                        village_text = area_data_cleaned[start_pos:end_pos]

                        # 先提取自由學區
                        free_zones = extract_free_zone_info(village_text)

                        # 🔥 如果有「以上X里」的共享自由學區，且當前村里沒有自己的自由學區資訊，
                        # 則使用共享的自由學區資訊
                        if shared_free_zones and not free_zones:
                            free_zones = shared_free_zones

                        # 將自由學區的鄰號記錄下來
                        free_zone_neighbors = set()
                        for fz in free_zones:
                            free_zone_neighbors.update(fz['neighbors'])

                            # 加入自由學區資訊
                            shared_schools = [s for s in fz['schools'] if s != current_school_name]
                            school_data[current_school_name][current_district][village]['free_zones'].append({
                                'neighbors': fz['neighbors'],
                                'shared_with': shared_schools
                            })

                        # 解析所有鄰號
                        all_neighbors = parse_neighbor_list(village_text)

                        # 普通學區 = 所有鄰號 - 自由學區鄰號
                        regular_neighbors = [n for n in all_neighbors if n not in free_zone_neighbors]

                        # 🔥 處理「全里」情況：
                        # 如果村里名稱後面沒有鄰號，且不是自由學區的一部分，則標記為「全里」
                        # 我們用特殊值 [-1] 表示「全里」
                        is_full_village = False

                        # 🔥 檢測「全里為XX學區」格式
                        # 例如：「新上里(全里為龍華國中學區，其中32鄰、33鄰、53鄰為龍華國中與明華國中自由學區)」
                        # 這種情況下，全里都是該學區，但部分鄰號同時是自由學區
                        if re.search(r'全里為.*學區', village_text):
                            is_full_village = True
                            regular_neighbors = [-1]

                        if not all_neighbors and not free_zones:
                            # 檢查這個村里是否只是列在列表中（如「文英里、文山里、文德里」）
                            # 這種情況表示全里都屬於該學區
                            is_full_village = True
                            regular_neighbors = [-1]  # 特殊標記表示「全里」

                        # 合併到已有的普通學區鄰號
                        existing = set(school_data[current_school_name][current_district][village]['regular'])
                        if is_full_village and -1 not in existing:
                            # 如果是全里，用 [-1] 標記
                            existing.add(-1)
                        else:
                            existing.update(regular_neighbors)
                        school_data[current_school_name][current_district][village]['regular'] = sorted(list(existing))

    return school_data


def save_to_json(data, output_filename, force=False):
    """儲存為 JSON 檔案，含驗證守門員（防止 buggy parser 覆蓋好資料）

    Args:
        data: 要存的資料 dict
        output_filename: 檔名
        force: True = 強制覆蓋（跳過守門員）；False = 啟用守門員（預設）

    Raises:
        RuntimeError: 守門員拒絕覆蓋時
    """
    # 🔥 使用 get_internal_dir() 取得 _internal 目錄
    internal_dir = get_internal_dir()
    output_path = os.path.join(internal_dir, output_filename)

    # 🔥 驗證守門員：比對舊版，「學校數」或「村里總數」任一指標不足 95% 就拒絕覆蓋
    # 避免 parser 對新格式 PDF 解析失敗時把好資料蓋掉
    # 村里總數是更敏感的指標（學校沒減多少但某區整個漏掉時會被抓到）
    def _count_villages(d):
        """計算總村里數：所有學校 × 所有區 × 該區的里"""
        if not isinstance(d, dict):
            return 0
        total = 0
        for school in d.values():
            if isinstance(school, dict):
                for district in school.values():
                    if isinstance(district, dict):
                        total += len(district)
        return total

    if not force and os.path.exists(output_path):
        try:
            with open(output_path, 'r', encoding='utf-8') as f:
                old_data = json.load(f)
            old_school_n = len(old_data) if isinstance(old_data, dict) else 0
            new_school_n = len(data) if isinstance(data, dict) else 0
            old_village_n = _count_villages(old_data)
            new_village_n = _count_villages(data)

            school_pct = (new_school_n / old_school_n * 100) if old_school_n else 100
            village_pct = (new_village_n / old_village_n * 100) if old_village_n else 100

            # 任一指標 < 95% 就擋
            if (old_school_n > 0 and new_school_n < old_school_n * 0.95) or \
               (old_village_n > 0 and new_village_n < old_village_n * 0.95):
                msg = (
                    f"\n⚠ ===== 驗證守門員：拒絕覆蓋 =====\n"
                    f"   舊版：{old_school_n} 校 / {old_village_n} 村里\n"
                    f"   新版：{new_school_n} 校 / {new_village_n} 村里\n"
                    f"   學校：{school_pct:.1f}%；村里：{village_pct:.1f}%（任一 <95% 就擋）\n"
                    f"   可能 PDF 格式變了 parser 解析失敗，已保留舊版 JSON\n"
                    f"   檔案: {output_path}"
                )
                print(msg)
                raise RuntimeError(
                    f"驗證未通過：{output_filename} 新版 {new_school_n}校/{new_village_n}里"
                    f" vs 舊版 {old_school_n}校/{old_village_n}里"
                    f"（{school_pct:.0f}%/{village_pct:.0f}%），已拒絕覆蓋以保護資料"
                )
        except json.JSONDecodeError:
            print(f"⚠ 舊 JSON 損毀，直接覆蓋")
        except (FileNotFoundError, OSError):
            pass  # 不影響

    with open(output_path, 'w', encoding='utf-8') as f:
        json.dump(data, f, ensure_ascii=False, indent=2)
    print(f"\n已儲存至: {output_path}")

    # 🔥 返回完整路徑供呼叫者使用
    return output_path


def print_summary(data, school_type):
    """輸出摘要"""
    print(f"\n{'=' * 80}")
    print(f"{school_type} 學區資料摘要")
    print(f"{'=' * 80}")
    print(f"總共 {len(data)} 所學校")

    for school_name, districts in data.items():
        total_villages = sum(len(villages) for villages in districts.values())
        print(f"\n{school_name}:")
        for district, villages in districts.items():
            print(f"  {district}: {len(villages)} 個村里")
            for village, info in villages.items():
                regular_count = len(info['regular'])
                free_zone_count = sum(len(fz['neighbors']) for fz in info['free_zones'])
                total = regular_count + free_zone_count
                print(f"    - {village}: {total} 鄰 (普通: {regular_count}, 自由: {free_zone_count})")


if __name__ == "__main__":
    # 建立國中學區資料
    print("開始建立國中學區資料...")
    junior_data = build_school_district_data('001.國中學區來源檔.pdf', '國中')
    save_to_json(junior_data, '高雄市_國中學區資料.json')
    print_summary(junior_data, '國中')

    print("\n" + "=" * 80)
    print("\n開始建立國小學區資料...")
    elementary_data = build_school_district_data('001.國小學區來源檔.pdf', '國小')
    save_to_json(elementary_data, '高雄市_國小學區資料.json')
    print_summary(elementary_data, '國小')

    print("\n" + "=" * 80)
    print("完成！")
