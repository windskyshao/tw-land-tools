# -*- coding: utf-8 -*-
"""
屏東縣學區資料建構程式
從屏東縣學區PDF檔案建立結構化JSON資料
"""

import pdfplumber
import re
import json
import os
import sys
import io

# 🔥 設定 UTF-8 輸出，避免 Windows 控制台編碼錯誤
if sys.stdout is not None:
    try:
        sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
    except AttributeError:
        pass

# 🔥 引入 base_dir_helper 以取得正確的輸出路徑
try:
    from base_dir_helper import get_internal_dir
except ImportError:
    # 如果無法引入，使用當前目錄
    def get_internal_dir():
        return os.path.dirname(os.path.abspath(__file__))


def parse_neighbor_list(text):
    """
    解析鄰號列表，支援多種格式

    範例：
    - "3.4.5.8.9.10.15.16.17.18鄰" -> [3, 4, 5, 8, 9, 10, 15, 16, 17, 18]
    - "1-5鄰" -> [1, 2, 3, 4, 5]
    - "1、3、5鄰" -> [1, 3, 5]
    - "全部" -> 'all'
    """
    if not text or not isinstance(text, str):
        return []

    # 移除換行和多餘空白
    text = text.replace('\n', '').replace(' ', '').replace('　', '')

    # 如果包含「全部」或「全里」，返回 'all'
    if '全部' in text or '全里' in text or '全村' in text:
        return 'all'

    neighbors = set()

    # 格式1: 用句點分隔的鄰號 "3.4.5.8.9.10鄰"
    dot_pattern = r'(?:^|[^0-9])(\d+(?:\.\d+)+)鄰'
    dot_matches = re.findall(dot_pattern, text)
    for match in dot_matches:
        nums = [int(n) for n in match.split('.') if n]
        neighbors.update(nums)

    # 格式2: 範圍格式 "1-5鄰" 或 "7至10鄰"
    range_pattern = r'(\d+)\s*[-－至到]\s*(\d+)\s*鄰?'
    for match in re.finditer(range_pattern, text):
        start = int(match.group(1))
        end = int(match.group(2))
        neighbors.update(range(start, end + 1))

    # 格式3: 單獨的鄰號 "1鄰"、"20鄰"
    single_pattern = r'(?<![0-9.-])(\d+)\s*鄰'
    for match in re.finditer(single_pattern, text):
        num = int(match.group(1))
        neighbors.add(num)

    # 格式4: 頓號或逗號分隔 "1、3、5鄰"
    comma_pattern = r'(\d+)\s*[、，]'
    for match in re.finditer(comma_pattern, text):
        num = int(match.group(1))
        neighbors.add(num)

    if neighbors:
        return sorted(list(neighbors))
    else:
        return []


def extract_shared_schools(notes_text):
    """
    從備註中提取共同學區的學校名稱

    範例：
    - "◎與萬新國中為共同學區" -> ['萬新國中']
    - "◎與南榮國中、東新國中為共同學區" -> ['南榮國中', '東新國中']
    """
    if not notes_text or not isinstance(notes_text, str):
        return []

    shared_schools = []

    # 匹配「與XX國中/國小為共同學區」的模式
    pattern = r'與([^為]+)為共同學區'
    matches = re.findall(pattern, notes_text)

    for match in matches:
        # 提取學校名稱（可能有多個學校用頓號分隔）
        school_pattern = r'([^、，\s]+(?:國中|國小))'
        schools = re.findall(school_pattern, match)
        shared_schools.extend(schools)

    return shared_schools


def parse_village_data(area_text):
    """
    解析學區範圍文字，提取村里和鄰號

    範例輸入：
    "大同里、楠樹里3.4.5.8.9.10.15.16.17.18鄰(仁愛路以西為界)、金泉里6.7.8.9.10.11.12鄰"

    返回：
    {
        '大同里': 'all',
        '楠樹里': [3, 4, 5, 8, 9, 10, 15, 16, 17, 18],
        '金泉里': [6, 7, 8, 9, 10, 11, 12]
    }
    """
    if not area_text or not isinstance(area_text, str):
        return {}

    result = {}

    # 移除換行
    area_text = area_text.replace('\n', '')

    # 找出所有村里及其位置
    village_pattern = r'([^、，。：\s(（]+[里村])'
    village_matches = []

    for match in re.finditer(village_pattern, area_text):
        village_name = match.group(1)
        # 排除包含這些關鍵字的假村里名稱
        if '共同' not in village_name and '學區' not in village_name and '自由' not in village_name:
            village_matches.append({
                'name': village_name,
                'start': match.start(),
                'end': match.end()
            })

    # 對每個村里，提取從它開始到下一個村里或分隔符之間的文字
    for i, village_info in enumerate(village_matches):
        village = village_info['name']

        # 提取這個村里的文字段落
        start_pos = village_info['end']

        # 找下一個村里或分隔符的位置
        if i + 1 < len(village_matches):
            end_pos = village_matches[i + 1]['start']
        else:
            end_pos = len(area_text)

        # 提取村里後面的文字（鄰號部分）
        village_text = area_text[start_pos:end_pos]

        # 解析鄰號
        neighbors = parse_neighbor_list(village + village_text)

        # 如果沒有找到鄰號，表示全里
        if not neighbors:
            result[village] = 'all'
        else:
            result[village] = neighbors

    return result


def build_pingtung_school_district_data(pdf_path, school_type='國中'):
    """
    從屏東縣學區PDF建立學區資料結構

    參數:
        pdf_path: PDF檔案路徑
        school_type: '國中' 或 '國小'

    返回:
        school_data: {
            '學校名稱': {
                '鄉鎮區': {
                    '村里': {
                        'basic': [鄰號列表] 或 'all',  # 基本學區
                        'shared': [               # 共同學區
                            {
                                'neighbors': [鄰號列表] 或 'all',
                                'shared_with': ['其他學校1', '其他學校2']
                            }
                        ]
                    }
                }
            }
        }
    """
    school_data = {}
    current_township = None  # 當前鄉鎮

    print(f"開始解析屏東縣 {school_type} PDF: {pdf_path}")

    with pdfplumber.open(pdf_path) as pdf:
        for page_num, page in enumerate(pdf.pages, 1):
            print(f"  處理第 {page_num} 頁...")

            tables = page.extract_tables()

            if not tables:
                # 嘗試從整頁文字中提取鄉鎮名稱
                page_text = page.extract_text()
                if page_text:
                    # 尋找鄉鎮名稱（通常在頁面開頭）
                    township_pattern = r'^([^、\n]{2,6}[鄉鎮市區])'
                    township_match = re.search(township_pattern, page_text, re.MULTILINE)
                    if township_match:
                        current_township = township_match.group(1)
                        print(f"    發現鄉鎮: {current_township}")
                continue

            for table in tables:
                if not table:
                    continue

                for row_idx, row in enumerate(table):
                    if not row:
                        continue

                    # 跳過表頭
                    if len(row) >= 2 and row[0] and '學校名稱' in str(row[0]):
                        continue

                    # 檢查是否是鄉鎮標題行
                    # 格式: "XX市各國民小學學區範圍一覽表(大學區制)"
                    if len(row) >= 1 and row[0]:
                        cell = str(row[0]).strip()
                        township_title_pattern = r'([^、]{2,6}[鄉鎮市區])(?:各)?國民(?:小學|中學)學區範圍一覽表'
                        township_match = re.search(township_title_pattern, cell)
                        if township_match:
                            current_township = township_match.group(1)
                            print(f"    發現鄉鎮: {current_township}")
                            continue

                    # 標準格式: [學校名稱, 學區範圍（里、鄰）, 備註]
                    if len(row) < 2:
                        continue

                    school_name = None
                    area_data = None
                    notes = None

                    # 提取學校名稱（第0欄）
                    if row[0] and isinstance(row[0], str):
                        cell = row[0].strip()
                        if (school_type == '國中' and '國中' in cell) or \
                           (school_type == '國小' and ('國小' in cell or '國民小學' in cell)):
                            school_name = cell.replace('國民中學', '國中').replace('國民小學', '國小')

                    # 提取學區範圍（第1欄）
                    if len(row) > 1 and row[1]:
                        area_data = str(row[1]).strip()

                    # 提取備註（第2欄）
                    if len(row) > 2 and row[2]:
                        notes = str(row[2]).strip()

                    # 如果找到學校名稱和學區資料
                    if school_name and area_data:
                        # 初始化學校資料
                        if school_name not in school_data:
                            school_data[school_name] = {}
                            print(f"    發現學校: {school_name}")

                        # 使用當前鄉鎮作為行政區
                        if not current_township:
                            current_township = '未知鄉鎮'

                        if current_township not in school_data[school_name]:
                            school_data[school_name][current_township] = {}

                        # 解析村里和鄰號
                        villages = parse_village_data(area_data)

                        for village, neighbors in villages.items():
                            if village not in school_data[school_name][current_township]:
                                school_data[school_name][current_township][village] = {
                                    'basic': neighbors,
                                    'shared': []
                                }
                            else:
                                # 合併基本學區鄰號
                                existing = school_data[school_name][current_township][village]['basic']
                                if neighbors == 'all':
                                    school_data[school_name][current_township][village]['basic'] = 'all'
                                elif existing != 'all':
                                    if isinstance(existing, list) and isinstance(neighbors, list):
                                        combined = set(existing + neighbors)
                                        school_data[school_name][current_township][village]['basic'] = sorted(list(combined))

                        # 解析共同學區
                        if notes:
                            shared_schools = extract_shared_schools(notes)
                            if shared_schools:
                                # 將所有村里標記為共同學區
                                for village in villages.keys():
                                    if village in school_data[school_name][current_township]:
                                        school_data[school_name][current_township][village]['shared'].append({
                                            'neighbors': villages[village],
                                            'shared_with': shared_schools
                                        })

    print(f"完成！共解析 {len(school_data)} 所{school_type}")
    return school_data


def save_json(data, output_path):
    """儲存JSON檔案"""
    with open(output_path, 'w', encoding='utf-8') as f:
        json.dump(data, f, ensure_ascii=False, indent=2)
    print(f"已儲存至：{output_path}")
    return output_path


def print_summary(data, school_type):
    """輸出摘要"""
    print(f"\n{'=' * 80}")
    print(f"屏東縣 {school_type} 學區資料摘要")
    print(f"{'=' * 80}")
    print(f"總共 {len(data)} 所學校")

    for school_name, townships in data.items():
        total_villages = sum(len(villages) for villages in townships.values())
        print(f"\n{school_name}:")
        for township, villages in townships.items():
            print(f"  {township}: {len(villages)} 個村里")
            for village, info in villages.items():
                basic = info['basic']
                basic_count = len(basic) if isinstance(basic, list) else '全里'
                shared_count = sum(len(s.get('neighbors', [])) if isinstance(s.get('neighbors'), list) else 0
                                 for s in info.get('shared', []))
                print(f"    - {village}: 基本 {basic_count} 鄰", end='')
                if info.get('shared'):
                    shared_with = ', '.join(info['shared'][0].get('shared_with', []))
                    print(f", 共同學區(與{shared_with})", end='')
                print()


def main():
    """主程式"""
    # 取得當前目錄（PDF檔案所在位置）
    current_dir = os.path.dirname(os.path.abspath(__file__))

    # 取得 _internal 目錄（JSON檔案輸出位置）
    internal_dir = get_internal_dir()

    # PDF檔案路徑（屏東縣）
    elementary_pdf = os.path.join(current_dir, '屏東縣 114 學年度各國民小學基本學區範圍一覽表.pdf')
    junior_pdf = os.path.join(current_dir, '屏東縣 114 學年度各國民中學基本學區範圍一覽表.pdf')

    # 檢查檔案是否存在
    if not os.path.exists(elementary_pdf):
        print(f"錯誤：找不到國小學區PDF檔案：{elementary_pdf}")
        print(f"請確認檔案是否存在於：{current_dir}")
        return

    if not os.path.exists(junior_pdf):
        print(f"錯誤：找不到國中學區PDF檔案：{junior_pdf}")
        print(f"請確認檔案是否存在於：{current_dir}")
        return

    # 解析國小學區
    print("\n" + "="*80)
    print("解析屏東縣國小學區")
    print("="*80)
    elementary_data = build_pingtung_school_district_data(elementary_pdf, '國小')
    elementary_output = os.path.join(internal_dir, '屏東縣國小學區資料.json')
    save_json(elementary_data, elementary_output)
    print_summary(elementary_data, '國小')

    # 解析國中學區
    print("\n" + "="*80)
    print("解析屏東縣國中學區")
    print("="*80)
    junior_data = build_pingtung_school_district_data(junior_pdf, '國中')
    junior_output = os.path.join(internal_dir, '屏東縣國中學區資料.json')
    save_json(junior_data, junior_output)
    print_summary(junior_data, '國中')

    print("\n" + "="*80)
    print("屏東縣學區資料轉換完成！")
    print("="*80)


if __name__ == '__main__':
    main()
