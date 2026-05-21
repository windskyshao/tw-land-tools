# -*- coding: utf-8 -*-
"""
學區查詢 - 使用預先建立的 JSON 資料
"""
import sys
import io
# 只在有 stdout.buffer 時才重新包裝（避免 GUI 環境中出錯）
if sys.stdout and hasattr(sys.stdout, 'buffer'):
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

import json
import os
from base_dir_helper import get_internal_dir


class SchoolDistrictLookup:
    """學區查詢類別"""

    def __init__(self, city='高雄市', data_dir=None):
        """
        初始化

        Args:
            city: 縣市名稱 ('高雄市', '台南市', '屏東縣')
            data_dir: JSON 資料檔案所在目錄，預設為 _internal 目錄
        """
        if data_dir is None:
            data_dir = get_internal_dir()

        self.city = city

        # 根據縣市選擇對應的 JSON 檔案
        if city == '高雄市':
            elementary_file = '高雄市_國小學區資料.json'
            junior_file = '高雄市_國中學區資料.json'
        elif city == '台南市':
            elementary_file = '台南市國小學區資料.json'
            junior_file = '台南市國中學區資料.json'
        elif city == '屏東縣':
            elementary_file = '屏東縣國小學區資料.json'
            junior_file = '屏東縣國中學區資料.json'
        else:
            print(f"警告：不支援的縣市 {city}，使用高雄市資料")
            elementary_file = '高雄市_國小學區資料.json'
            junior_file = '高雄市_國中學區資料.json'

        # 載入資料
        self.elementary_data = self._load_json(os.path.join(data_dir, elementary_file))
        self.junior_data = self._load_json(os.path.join(data_dir, junior_file))

    def _load_json(self, file_path):
        """載入 JSON 檔案"""
        try:
            with open(file_path, 'r', encoding='utf-8') as f:
                return json.load(f)
        except FileNotFoundError:
            print(f"警告：找不到檔案 {file_path}")
            return {}
        except json.JSONDecodeError as e:
            print(f"警告：JSON 格式錯誤 {file_path}: {e}")
            return {}

    def query(self, village, neighbor=None, district=None, school_type='國中'):
        """
        查詢學區

        Args:
            village: 村里名稱 (例如: "福山里")
            neighbor: 鄰號 (可選，數字或字串)
            district: 行政區 (可選，例如: "左營區")
            school_type: '國中' 或 '國小'

        Returns:
            list: 學校名稱列表
        """
        # 選擇資料
        data = self.junior_data if school_type == '國中' else self.elementary_data

        # 轉換鄰號為整數
        target_neighbor = int(neighbor) if neighbor else None

        results = []

        # 遍歷所有學校
        for school_name, districts in data.items():
            # 如果指定了行政區，只檢查該行政區
            if district:
                if district not in districts:
                    continue
                district_list = {district: districts[district]}
            else:
                district_list = districts

            # 檢查每個行政區
            for dist_name, villages in district_list.items():
                # 檢查村里是否存在
                if village not in villages:
                    continue

                village_data = villages[village]

                # 如果沒有指定鄰號，只要村里符合就加入
                if target_neighbor is None:
                    if school_name not in results:
                        results.append(school_name)
                    continue

                # 🔥 支援兩種格式：
                # 格式1（高雄市/台南市）：{'regular': [...], 'free_zones': [...]}
                # 格式2（屏東縣）：{'basic': [...] 或 'all', 'shared': [...]}

                # 判斷格式並提取鄰號列表
                if 'regular' in village_data:
                    # 格式1：高雄市/台南市
                    regular_neighbors = village_data.get('regular', [])
                    free_zones = village_data.get('free_zones', [])

                    # 🔥 如果 regular 包含 -1，表示「全里」都屬於該學區
                    if -1 in regular_neighbors:
                        if school_name not in results:
                            results.append(school_name)
                        continue

                    # 如果 regular 和 free_zones 都為空，視為所有鄰號都屬於該學區
                    if not regular_neighbors and not free_zones:
                        if school_name not in results:
                            results.append(school_name)
                        continue

                    # 檢查鄰號是否在普通學區
                    if target_neighbor in regular_neighbors:
                        if school_name not in results:
                            results.append(school_name)
                        continue

                    # 檢查鄰號是否在自由學區
                    for free_zone in free_zones:
                        fz_neighbors = free_zone['neighbors']
                        # 🔥 -1 代表「全里」，視為任何鄰號都匹配（與 regular 邏輯一致）
                        if -1 in fz_neighbors or target_neighbor in fz_neighbors:
                            if school_name not in results:
                                results.append(school_name)
                            for shared_school in free_zone['shared_with']:
                                if shared_school not in results:
                                    results.append(shared_school)

                elif 'basic' in village_data:
                    # 格式2：屏東縣
                    basic = village_data.get('basic', [])
                    shared = village_data.get('shared', [])

                    # basic 可能是 "all" 或鄰號列表
                    if basic == "all":
                        if school_name not in results:
                            results.append(school_name)
                        continue

                    # 檢查鄰號是否在基本學區
                    if isinstance(basic, list) and target_neighbor in basic:
                        if school_name not in results:
                            results.append(school_name)
                        continue

                    # 檢查鄰號是否在共享學區
                    for shared_zone in shared:
                        if target_neighbor in shared_zone.get('neighbors', []):
                            if school_name not in results:
                                results.append(school_name)
                            for shared_school in shared_zone.get('shared_with', []):
                                if shared_school not in results:
                                    results.append(shared_school)

        return results

    def query_schools(self, village, neighbor=None, district=None, log_callback=None):
        """
        學區查詢 (同時查詢國小和國中)

        Args:
            village: 村里名稱
            neighbor: 鄰號 (可選)
            district: 行政區 (可選)
            log_callback: 日誌回調函數 (可選)

        Returns:
            dict: {'國小': [...], '國中': [...]}
        """
        log = log_callback or print

        log(f"[步驟 7] 查詢{self.city}學區...")
        log(f"  村里：{village}")
        if neighbor:
            log(f"  鄰：第{neighbor}鄰")
        if district:
            log(f"  行政區：{district}")

        result = {
            '國小': [],
            '國中': []
        }

        # 查詢國小
        log(f"  ⏳ 正在查詢國小學區...")
        elementary_schools = self.query(village, neighbor, district, '國小')
        result['國小'] = elementary_schools
        if elementary_schools:
            for school in elementary_schools:
                log(f"  ✓ 找到國小：{school}")

        # 查詢國中
        log(f"  ⏳ 正在查詢國中學區...")
        junior_schools = self.query(village, neighbor, district, '國中')
        result['國中'] = junior_schools
        if junior_schools:
            for school in junior_schools:
                log(f"  ✓ 找到國中：{school}")

        # 輸出結果
        if result['國小'] or result['國中']:
            log(f"\n✓ 學區查詢完成：")
            if result['國小']:
                formatted = self._format_school_list(result['國小'], '國小')
                log(f"  國小：{formatted}")
            if result['國中']:
                formatted = self._format_school_list(result['國中'], '國中')
                log(f"  國中：{formatted}")
        else:
            log(f"\n⚠ 未查詢到學區資料")

        return result

    def _format_school_list(self, schools, suffix='國中'):
        """格式化學校列表"""
        if not schools:
            return ''

        # 移除後綴以便顯示
        names = [s.replace(suffix, '') for s in schools]

        if len(names) == 1:
            return names[0] + suffix
        else:
            return '、'.join(names) + suffix


# 全域實例字典（支援多縣市）
_lookup_instances = {}


def get_lookup(city='高雄市'):
    """
    獲取學區查詢實例

    Args:
        city: 縣市名稱 ('高雄市', '台南市', '屏東縣')

    Returns:
        SchoolDistrictLookup: 學區查詢實例
    """
    global _lookup_instances
    if city not in _lookup_instances:
        _lookup_instances[city] = SchoolDistrictLookup(city=city)
    return _lookup_instances[city]


def query_kaohsiung_schools(village, neighbor=None, district=None, log_callback=None):
    """
    便捷函數：查詢高雄市學區

    Args:
        village: 村里名稱
        neighbor: 鄰號 (可選)
        district: 行政區 (可選)
        log_callback: 日誌回調函數 (可選)

    Returns:
        dict: {'國小': [...], '國中': [...]}
    """
    lookup = get_lookup('高雄市')
    return lookup.query_schools(village, neighbor, district, log_callback)


def query_tainan_schools(village, neighbor=None, district=None, log_callback=None):
    """
    便捷函數：查詢台南市學區

    Args:
        village: 村里名稱
        neighbor: 鄰號 (可選)
        district: 行政區 (可選)
        log_callback: 日誌回調函數 (可選)

    Returns:
        dict: {'國小': [...], '國中': [...]}
    """
    lookup = get_lookup('台南市')
    return lookup.query_schools(village, neighbor, district, log_callback)


def query_pingtung_schools(village, neighbor=None, district=None, log_callback=None):
    """
    便捷函數：查詢屏東縣學區

    Args:
        village: 村里名稱
        neighbor: 鄰號 (可選)
        district: 行政區 (可選)
        log_callback: 日誌回調函數 (可選)

    Returns:
        dict: {'國小': [...], '國中': [...]}
    """
    lookup = get_lookup('屏東縣')
    return lookup.query_schools(village, neighbor, district, log_callback)


def query_schools(city, village, neighbor=None, district=None, log_callback=None):
    """
    通用便捷函數：查詢指定縣市的學區

    Args:
        city: 縣市名稱 ('高雄市', '台南市', '屏東縣')
        village: 村里名稱
        neighbor: 鄰號 (可選)
        district: 行政區 (可選)
        log_callback: 日誌回調函數 (可選)

    Returns:
        dict: {'國小': [...], '國中': [...]}
    """
    lookup = get_lookup(city)
    return lookup.query_schools(village, neighbor, district, log_callback)


if __name__ == "__main__":
    print("\n" + "=" * 80)
    print("測試案例 1: 福山里1鄰（預期：福山國中）")
    print("=" * 80)

    result1 = query_kaohsiung_schools("福山里", neighbor="1", district="左營區")
    print(f"\n結果：國中 = {result1['國中']}")
    if '福山國中' in result1['國中']:
        print("✓✓✓ 測試 1 通過！")
    else:
        print("✗✗✗ 測試 1 失敗！")

    print("\n" + "=" * 80)
    print("測試案例 2: 福山里65鄰（預期：文府國小、文府國中、福山國中）")
    print("=" * 80)

    result2 = query_kaohsiung_schools("福山里", neighbor="65", district="左營區")
    print(f"\n結果：")
    print(f"  國小 = {result2['國小']}")
    print(f"  國中 = {result2['國中']}")

    # 驗證結果
    expected_elementary = ['文府國小']
    expected_junior = {'福山國中', '文府國中'}

    if (result2['國小'] == expected_elementary and
        set(result2['國中']) == expected_junior):
        print("✓✓✓ 測試 2 通過！")
    else:
        print("✗✗✗ 測試 2 失敗！")
        if result2['國小'] != expected_elementary:
            print(f"  國小預期：{expected_elementary}，實際：{result2['國小']}")
        if set(result2['國中']) != expected_junior:
            print(f"  國中預期：{expected_junior}，實際：{set(result2['國中'])}")
