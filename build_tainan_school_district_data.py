"""
台南市學區資料建構程式
從台南市學區PDF檔案建立結構化JSON資料
"""

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
    解析鄰號列表，支援多種格式
    範例：
    - "1-5 鄰" -> [1, 2, 3, 4, 5]
    - "7-12 鄰" -> [7, 8, 9, 10, 11, 12]
    - "1 至 5 鄰" -> [1, 2, 3, 4, 5]
    - "1、3、5 鄰" -> [1, 3, 5]
    - "1 鄰" -> [1]
    """
    if not text or not isinstance(text, str):
        return []

    neighbors = []

    # 移除全形數字，轉為半形
    text = text.replace('０', '0').replace('１', '1').replace('２', '2')
    text = text.replace('３', '3').replace('４', '4').replace('５', '5')
    text = text.replace('６', '6').replace('７', '7').replace('８', '8')
    text = text.replace('９', '9')

    # 格式1: "7-12 鄰" 或 "65鄰-67鄰"
    range_pattern = r'(\d+)\s*[-－至到]\s*(\d+)\s*鄰?'
    for match in re.finditer(range_pattern, text):
        start = int(match.group(1))
        end = int(match.group(2))
        neighbors.extend(range(start, end + 1))

    # 格式2: 單獨的鄰號，如 "1 鄰"、"20鄰"
    single_pattern = r'(\d+)\s*鄰'
    for match in re.finditer(single_pattern, text):
        num = int(match.group(1))
        if num not in neighbors:
            neighbors.append(num)

    # 格式3: 頓號分隔，如 "1、3、5"
    comma_nums = re.findall(r'(\d+)\s*[、，]', text)
    for num_str in comma_nums:
        num = int(num_str)
        if num not in neighbors:
            neighbors.append(num)

    return sorted(list(set(neighbors)))


def parse_district_village_neighbors(area_text):
    """
    解析台南市格式的學區資料
    格式：區名：里名 鄰號範圍，里名 鄰號範圍

    範例輸入：
    "東區：泉南里，東門里 7-12 鄰，大同里 5-8、12-13 鄰"

    返回：
    {
        '東區': {
            '泉南里': [全部鄰號],  # 沒指定鄰號表示全里
            '東門里': [7, 8, 9, 10, 11, 12],
            '大同里': [5, 6, 7, 8, 12, 13]
        }
    }
    """
    result = {}

    if not area_text:
        return result

    # 將全形符號轉半形
    area_text = str(area_text).replace('，', '，').replace('：', '：')

    # 分割不同行政區的資料（用中文換行或全形逗號分隔大區塊）
    # 先按行分割
    lines = area_text.split('\n')

    for line in lines:
        line = line.strip()
        if not line:
            continue

        # 嘗試匹配格式：區名：里名資料
        district_match = re.match(r'([^：]+)[：:](.+)', line)
        if not district_match:
            continue

        district_name = district_match.group(1).strip()
        village_data = district_match.group(2).strip()

        if district_name not in result:
            result[district_name] = {}

        # 解析里和鄰號
        # 將村里資料以頓號或逗號分割
        village_parts = re.split(r'[，、]', village_data)

        current_village = None
        current_neighbors = []

        for part in village_parts:
            part = part.strip()
            if not part:
                continue

            # 檢查是否包含"里"字（表示新的村里開始）
            if '里' in part:
                # 儲存前一個村里的資料
                if current_village and current_neighbors:
                    if current_village not in result[district_name]:
                        result[district_name][current_village] = []
                    result[district_name][current_village].extend(current_neighbors)
                elif current_village:
                    # 沒有鄰號表示全里
                    result[district_name][current_village] = 'all'

                # 提取村里名稱和鄰號
                village_match = re.match(r'(.+里)\s*(.*)', part)
                if village_match:
                    current_village = village_match.group(1).strip()
                    neighbor_text = village_match.group(2).strip()
                    current_neighbors = parse_neighbor_list(neighbor_text) if neighbor_text else []

                    # 如果這是最後一個，直接儲存
                    if part == village_parts[-1]:
                        if current_neighbors:
                            if current_village not in result[district_name]:
                                result[district_name][current_village] = []
                            result[district_name][current_village].extend(current_neighbors)
                        else:
                            result[district_name][current_village] = 'all'
            else:
                # 這部分只有鄰號，屬於前一個村里
                if current_village:
                    neighbors = parse_neighbor_list(part)
                    current_neighbors.extend(neighbors)

                    # 如果這是最後一個，儲存
                    if part == village_parts[-1]:
                        if current_village not in result[district_name]:
                            result[district_name][current_village] = []
                        result[district_name][current_village].extend(current_neighbors)

    # 去重並排序
    for district in result:
        for village in result[district]:
            if result[district][village] != 'all':
                result[district][village] = sorted(list(set(result[district][village])))

    return result


def build_tainan_school_district_data(pdf_path, school_type='國中'):
    """
    從台南市學區PDF建立學區資料結構

    參數:
        pdf_path: PDF檔案路徑
        school_type: '國中' 或 '國小'

    返回:
        school_data: {
            '學校名稱': {
                '行政區': {
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
    import pandas as pd

    school_data = {}

    print(f"開始解析 {school_type} PDF: {pdf_path}")

    with pdfplumber.open(pdf_path) as pdf:
        for page_num, page in enumerate(pdf.pages, 1):
            print(f"處理第 {page_num} 頁...")
            tables = page.extract_tables()

            if not tables:
                continue

            for table in tables:
                # 跳過表頭
                for row in table[1:]:  # 跳過第一行表頭
                    if not row or len(row) < 4:
                        continue

                    # 台南市格式：編號、區別、學校名稱、基本學區、共同學區
                    try:
                        school_num = row[0]
                        district_col = row[1]
                        school_name = row[2]
                        basic_area = row[3] if len(row) > 3 else None
                        shared_area = row[4] if len(row) > 4 else None
                    except IndexError:
                        continue

                    # 跳過空行或表頭行
                    if not school_name or school_name in ['學校', '名稱', '']:
                        continue

                    # 清理學校名稱（移除空白和換行）
                    school_name = str(school_name).strip()
                    school_name = school_name.replace('\n', '').replace(' ', '').replace('　', '')
                    if not school_name or school_name in ['學校名稱', '']:
                        continue

                    # 確保學校名稱以"國中"或"國小"結尾
                    if school_type == '國中' and '國中' not in school_name:
                        continue
                    if school_type == '國小' and '國小' not in school_name:
                        continue

                    # 初始化學校資料
                    if school_name not in school_data:
                        school_data[school_name] = {}

                    # 解析基本學區
                    if basic_area:
                        basic_districts = parse_district_village_neighbors(basic_area)
                        for dist, villages in basic_districts.items():
                            if dist not in school_data[school_name]:
                                school_data[school_name][dist] = {}

                            for village, neighbors in villages.items():
                                if village not in school_data[school_name][dist]:
                                    school_data[school_name][dist][village] = {
                                        'basic': neighbors,
                                        'shared': []
                                    }
                                else:
                                    # 合併基本學區鄰號
                                    if neighbors == 'all':
                                        school_data[school_name][dist][village]['basic'] = 'all'
                                    elif school_data[school_name][dist][village]['basic'] != 'all':
                                        if isinstance(school_data[school_name][dist][village]['basic'], list):
                                            school_data[school_name][dist][village]['basic'].extend(neighbors)
                                            school_data[school_name][dist][village]['basic'] = sorted(list(set(school_data[school_name][dist][village]['basic'])))

                    # 解析共同學區
                    if shared_area:
                        # 共同學區格式較複雜，需要解析出：
                        # 1. 哪些區里鄰號
                        # 2. 與哪些學校共同

                        # 使用正則表達式分割共同學區項目
                        # 格式：區名：里名 鄰號(與XX學校共同學區)
                        shared_items = re.split(r'\n|(?=[^）]+：)', str(shared_area))

                        for item in shared_items:
                            item = item.strip()
                            if not item:
                                continue

                            # 提取共同學校
                            shared_schools = []
                            school_pattern = r'([^與（）\s]+?(?:國中|國小|實小|高中))'
                            shared_schools = re.findall(school_pattern, item)

                            # 移除當前學校自己
                            shared_schools = [s for s in shared_schools if s != school_name and s.strip()]

                            # 解析區里鄰號
                            districts = parse_district_village_neighbors(item)

                            for dist, villages in districts.items():
                                if dist not in school_data[school_name]:
                                    school_data[school_name][dist] = {}

                                for village, neighbors in villages.items():
                                    if village not in school_data[school_name][dist]:
                                        school_data[school_name][dist][village] = {
                                            'basic': [],
                                            'shared': []
                                        }

                                    # 加入共同學區資料
                                    school_data[school_name][dist][village]['shared'].append({
                                        'neighbors': neighbors,
                                        'shared_with': shared_schools
                                    })

    print(f"完成！共解析 {len(school_data)} 所{school_type}")
    return school_data


def save_json(data, output_path):
    """儲存JSON檔案"""
    with open(output_path, 'w', encoding='utf-8') as f:
        json.dump(data, f, ensure_ascii=False, indent=2)
    print(f"已儲存至：{output_path}")
    # 🔥 返回完整路徑供呼叫者使用
    return output_path


def main():
    """主程式"""
    # 🔥 取得當前目錄（PDF檔案所在位置）
    current_dir = os.path.dirname(os.path.abspath(__file__))

    # 🔥 取得 _internal 目錄（JSON檔案輸出位置）
    internal_dir = get_internal_dir()

    # PDF檔案路徑（台南市）
    elementary_pdf = os.path.join(current_dir, '臺南市114學年度國民小學學區一覽表.pdf')
    junior_pdf = os.path.join(current_dir, '臺南市114學年度國民中學學區一覽表.pdf')

    # 檢查檔案是否存在
    if not os.path.exists(elementary_pdf):
        print(f"錯誤：找不到國小學區PDF檔案：{elementary_pdf}")
        return

    if not os.path.exists(junior_pdf):
        print(f"錯誤：找不到國中學區PDF檔案：{junior_pdf}")
        return

    # 解析國小學區
    print("\n" + "="*60)
    print("解析台南市國小學區")
    print("="*60)
    elementary_data = build_tainan_school_district_data(elementary_pdf, '國小')
    elementary_output = os.path.join(internal_dir, '台南市國小學區資料.json')  # 🔥 使用 internal_dir
    save_json(elementary_data, elementary_output)

    # 解析國中學區
    print("\n" + "="*60)
    print("解析台南市國中學區")
    print("="*60)
    junior_data = build_tainan_school_district_data(junior_pdf, '國中')
    junior_output = os.path.join(internal_dir, '台南市國中學區資料.json')  # 🔥 使用 internal_dir
    save_json(junior_data, junior_output)

    print("\n" + "="*60)
    print("台南市學區資料轉換完成！")
    print("="*60)


if __name__ == '__main__':
    # 需要 pandas 來處理某些資料
    import pandas as pd
    main()
