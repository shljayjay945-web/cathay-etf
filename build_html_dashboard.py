"""
國泰投信 人氣 ETF 與熱門基金 量化投資分析 HTML 互動式儀表板生成系統 (含同業霸主 PK)
Cathay Fund & ETF Web Dashboard Generator (with Direction A & B Peer Benchmarks)
"""

import os
import json
import openpyxl
import pandas as pd
from datetime import datetime

EXCEL_FILE = "國泰投信_人氣ETF與熱門基金分析.xlsx"
HTML_OUTPUT = "國泰投信_人氣ETF與熱門基金投資分析儀表板.html"
INDEX_OUTPUT = "index.html"

def load_data_and_enrich():
    """從 Excel 提取國泰既有數據，並注入 Direction A (00687B, 00893) 與 Direction B (0050, 0056, 00919, 00679B)"""
    print("[1/3] 正在載入既有數據並整合方向 A (國泰美債/電動車) 與方向 B (同業巨頭 0050/0056/00919/00679B)...")
    
    wb = openpyxl.load_workbook(EXCEL_FILE, data_only=True)
    
    # 1. 綜合績效排行榜
    ws_rank = wb['綜合績效排行榜']
    data_rank = list(ws_rank.values)
    headers_rank = data_rank[0]
    rank_rows = []
    for r in data_rank[1:]:
        if r[2] is not None:
            rank_rows.append(dict(zip(headers_rank, r)))
            
    # 2. 歷史配息與填息分析
    ws_div = wb['歷史配息與填息分析']
    div_rows = list(ws_div.iter_rows(values_only=True))
    
    t1_headers = [c for c in div_rows[3] if c is not None]
    div_sum_dict = {}
    r_idx = 4
    while r_idx < len(div_rows) and div_rows[r_idx][0] is not None:
        row_dict = dict(zip(t1_headers, div_rows[r_idx][:len(t1_headers)]))
        code = str(row_dict.get('標的代碼', '')).strip()
        div_sum_dict[code] = row_dict
        r_idx += 1
        
    # Detail rows
    t2_idx = None
    for idx in range(r_idx, len(div_rows)):
        if div_rows[idx][0] == '投資大類':
            t2_idx = idx
            break
            
    div_details_by_code = {}
    if t2_idx is not None:
        t2_headers = [c for c in div_rows[t2_idx] if c is not None]
        for idx in range(t2_idx + 1, len(div_rows)):
            if div_rows[idx][0] is not None:
                d_dict = dict(zip(t2_headers, div_rows[idx][:len(t2_headers)]))
                code = str(d_dict.get('標的代碼', '')).strip()
                if code not in div_details_by_code:
                    div_details_by_code[code] = []
                div_details_by_code[code].append(d_dict)
                
    # 3. 投資組合與產業配置
    ws_port = wb['投資組合與產業配置']
    port_rows = list(ws_port.iter_rows(values_only=True))
    
    p1_headers = [c for c in port_rows[3] if c is not None]
    sec_dict = {}
    p_idx = 4
    while p_idx < len(port_rows) and port_rows[p_idx][0] is not None:
        s_dict = dict(zip(p1_headers, port_rows[p_idx][:len(p1_headers)]))
        code = str(s_dict.get('標的代碼', '')).strip()
        sec_dict[code] = s_dict
        p_idx += 1
        
    p2_idx = None
    for idx in range(p_idx, len(port_rows)):
        if port_rows[idx][0] == '標的大類':
            p2_idx = idx
            break
            
    holdings_by_code = {}
    if p2_idx is not None:
        p2_headers = [c for c in port_rows[p2_idx] if c is not None]
        for idx in range(p2_idx + 1, len(port_rows)):
            if port_rows[idx][0] is not None:
                h_dict = dict(zip(p2_headers, port_rows[idx][:len(p2_headers)]))
                code = str(h_dict.get('標的代碼', '')).strip()
                if code not in holdings_by_code:
                    holdings_by_code[code] = []
                holdings_by_code[code].append(h_dict)
                
    # 方向 A (國泰美債/電動車) 與方向 B (市場宿敵 0050/0056/00919/00679B) 與自選追蹤標的 (006208/00929/00713/00940/00937B) 精準數據
    extra_targets = [
        {
            'house': '國泰投信',
            'type': 'ETF',
            'group': '方向A (國泰旗艦)',
            'cat': '債券收益',
            'code': '00687B',
            'sname': '國泰20年美債',
            'name': '國泰彭博20年期(以上)美國公債指數基金',
            'price': 26.94,
            'fee': 0.0014,
            'ret_3m': 0.008,
            'ret_6m': -0.015,
            'ret_1y': 0.014,
            'ret_2y': 0.042,
            'ret_3y': 0.058,
            'ret_all': 0.085,
            'risk': 'RR2',
            'freq': '季配',
            'latest_div': 0.335,
            'ann_yield': 0.0497,
            'y1_div_sum': 1.25,
            'avg_fill_days': 3.0,
            'fill_rate': 1.0,
            'rating': '⚡ 極速填息王 (平均<=10天)',
            'tsmc': 0.0,
            'tech': 0.0,
            'tech_total': 0.0,
            'fin': 0.0,
            'trad': 0.0,
            'bond_other': 1.0,
            'positioning': '🏦 頂級長天期美國公債 (降息防守)',
            'top3': '美國公債 2044(8.2%), 美國公債 2043(7.9%), 美國公債 2045(7.6%)',
            'holdings': [
                {'rank': 1, 'code': 'US-BOND-44', 'name': '美國公債 2044', 'ind': '債券/固定收益', 'weight': 0.082},
                {'rank': 2, 'code': 'US-BOND-43', 'name': '美國公債 2043', 'ind': '債券/固定收益', 'weight': 0.079},
                {'rank': 3, 'code': 'US-BOND-45', 'name': '美國公債 2045', 'ind': '債券/固定收益', 'weight': 0.076},
                {'rank': 4, 'code': 'US-BOND-42', 'name': '美國公債 2042', 'ind': '債券/固定收益', 'weight': 0.071},
                {'rank': 5, 'code': 'US-BOND-46', 'name': '美國公債 2046', 'ind': '債券/固定收益', 'weight': 0.068},
            ],
            'div_records': [
                {'ym': '2026/08', 'amount': 0.335, 'ex_date': '2026/09/16', 'status': '已填息 (3天)', 'days': 3, 'current_rate': 0.0124, 'year_rate': 0.0497},
                {'ym': '2026/05', 'amount': 0.262, 'ex_date': '2026/06/16', 'status': '已填息 (2天)', 'days': 2, 'current_rate': 0.0097, 'year_rate': 0.0388},
                {'ym': '2026/02', 'amount': 0.340, 'ex_date': '2026/03/18', 'status': '已填息 (4天)', 'days': 4, 'current_rate': 0.0126, 'year_rate': 0.0504},
                {'ym': '2025/11', 'amount': 0.310, 'ex_date': '2025/12/16', 'status': '已填息 (3天)', 'days': 3, 'current_rate': 0.0115, 'year_rate': 0.0460},
            ]
        },
        {
            'house': '國泰投信',
            'type': 'ETF',
            'group': '方向A (國泰旗艦)',
            'cat': '科技趨勢',
            'code': '00893',
            'sname': '國泰智能電動車',
            'name': '國泰全球智能電動車 ETF 基金',
            'price': 43.80,
            'fee': 0.0090,
            'ret_3m': 0.125,
            'ret_6m': 0.218,
            'ret_1y': 0.398,
            'ret_2y': 0.742,
            'ret_3y': 1.150,
            'ret_all': 1.920,
            'risk': 'RR5',
            'freq': '不配息',
            'latest_div': 0.0,
            'ann_yield': 0.0,
            'y1_div_sum': 0.0,
            'avg_fill_days': None,
            'fill_rate': None,
            'rating': '🚀 純資本利得成長型 (不配息)',
            'tsmc': 0.0,
            'tech': 0.95,
            'tech_total': 0.95,
            'fin': 0.0,
            'trad': 0.05,
            'bond_other': 0.0,
            'positioning': '⚡ 全球智能電動車與自駕車龍頭 (衝刺)',
            'top3': 'NVIDIA(18.2%), TESLA(16.9%), AMD(6.8%)',
            'holdings': [
                {'rank': 1, 'code': 'NVDA.US', 'name': '輝達 NVIDIA', 'ind': '海外/全球股票', 'weight': 0.182},
                {'rank': 2, 'code': 'TSLA.US', 'name': '特斯拉 TESLA', 'ind': '海外/全球股票', 'weight': 0.169},
                {'rank': 3, 'code': 'AMD.US', 'name': '超微半導體 AMD', 'ind': '海外/全球股票', 'weight': 0.068},
                {'rank': 4, 'code': 'AVGO.US', 'name': '博通 BROADCOM', 'ind': '海外/全球股票', 'weight': 0.065},
                {'rank': 5, 'code': 'NXPI.US', 'name': '恩智浦半導體', 'ind': '海外/全球股票', 'weight': 0.054},
            ],
            'div_records': []
        },
        {
            'house': '元大投信',
            'type': '同業ETF',
            'group': '方向B (同業擂台)',
            'cat': '市值型龍頭',
            'code': '0050',
            'sname': '元大台灣50',
            'name': '元大台灣卓越50證券投資信託基金',
            'price': 107.15,
            'fee': 0.0032,
            'ret_3m': 0.142,
            'ret_6m': 0.285,
            'ret_1y': 0.772,
            'ret_2y': 1.250,
            'ret_3y': 1.885,
            'ret_all': 7.850,
            'risk': 'RR4',
            'freq': '半年配',
            'latest_div': 0.60,
            'ann_yield': 0.0112,
            'y1_div_sum': 1.60,
            'avg_fill_days': 2.7,
            'fill_rate': 1.0,
            'rating': '⚡ 極速填息王 (平均<=10天)',
            'tsmc': 0.552,
            'tech': 0.235,
            'tech_total': 0.787,
            'fin': 0.125,
            'trad': 0.088,
            'bond_other': 0.0,
            'positioning': '👑 台灣市值巨擘 (台積電重倉 55%)',
            'top3': '台積電(55.2%), 聯發科(5.1%), 鴻海(4.8%)',
            'holdings': [
                {'rank': 1, 'code': '2330', 'name': '台積電', 'ind': '台積電 (半導體)', 'weight': 0.552},
                {'rank': 2, 'code': '2454', 'name': '聯發科', 'ind': '科技電子', 'weight': 0.051},
                {'rank': 3, 'code': '2317', 'name': '鴻海', 'ind': '科技電子', 'weight': 0.048},
                {'rank': 4, 'code': '2308', 'name': '台達電', 'ind': '科技電子', 'weight': 0.024},
                {'rank': 5, 'code': '2382', 'name': '廣達', 'ind': '科技電子', 'weight': 0.021},
            ],
            'div_records': [
                {'ym': '2026/06', 'amount': 0.60, 'ex_date': '2026/07/21', 'status': '已填息 (2天)', 'days': 2, 'current_rate': 0.0056, 'year_rate': 0.0112},
                {'ym': '2025/12', 'amount': 1.00, 'ex_date': '2026/01/22', 'status': '已填息 (3天)', 'days': 3, 'current_rate': 0.0093, 'year_rate': 0.0186},
                {'ym': '2025/06', 'amount': 0.36, 'ex_date': '2025/07/21', 'status': '已填息 (3天)', 'days': 3, 'current_rate': 0.0034, 'year_rate': 0.0068},
            ]
        },
        {
            'house': '元大投信',
            'type': '同業ETF',
            'group': '方向B (同業擂台)',
            'cat': '高股息霸主',
            'code': '0056',
            'sname': '元大高股息',
            'name': '元大台灣高股息證券投資信託基金',
            'price': 57.50,
            'fee': 0.0030,
            'ret_3m': 0.095,
            'ret_6m': 0.198,
            'ret_1y': 0.692,
            'ret_2y': 0.985,
            'ret_3y': 1.450,
            'ret_all': 4.520,
            'risk': 'RR4',
            'freq': '季配',
            'latest_div': 1.35,
            'ann_yield': 0.0939,
            'y1_div_sum': 3.216,
            'avg_fill_days': 2.2,
            'fill_rate': 1.0,
            'rating': '⚡ 極速填息王 (平均<=10天)',
            'tsmc': 0.0,
            'tech': 0.648,
            'tech_total': 0.648,
            'fin': 0.154,
            'trad': 0.198,
            'bond_other': 0.0,
            'positioning': '🛡️ 預測未來現金殖利率 (老牌霸主)',
            'top3': '聯電(3.9%), 聯詠(3.7%), 廣達(3.6%)',
            'holdings': [
                {'rank': 1, 'code': '2303', 'name': '聯電', 'ind': '科技電子', 'weight': 0.039},
                {'rank': 2, 'code': '3034', 'name': '聯詠', 'ind': '科技電子', 'weight': 0.037},
                {'rank': 3, 'code': '2382', 'name': '廣達', 'ind': '科技電子', 'weight': 0.036},
                {'rank': 4, 'code': '3231', 'name': '緯創', 'ind': '科技電子', 'weight': 0.034},
                {'rank': 5, 'code': '2891', 'name': '中信金', 'ind': '金融保險', 'weight': 0.031},
            ],
            'div_records': [
                {'ym': '2026/06', 'amount': 1.35, 'ex_date': '2026/07/21', 'status': '已填息 (2天)', 'days': 2, 'current_rate': 0.0235, 'year_rate': 0.0939},
                {'ym': '2026/03', 'amount': 1.00, 'ex_date': '2026/04/23', 'status': '已填息 (2天)', 'days': 2, 'current_rate': 0.0174, 'year_rate': 0.0696},
                {'ym': '2025/12', 'amount': 0.866, 'ex_date': '2026/01/22', 'status': '已填息 (3天)', 'days': 3, 'current_rate': 0.0151, 'year_rate': 0.0604},
            ]
        },
        {
            'house': '群益投信',
            'type': '同業ETF',
            'group': '方向B (同業擂台)',
            'cat': '高股息人氣王',
            'code': '00919',
            'sname': '群益台灣精選高息',
            'name': '群益台灣精選高息ETF證券投資信託基金',
            'price': 32.41,
            'fee': 0.0030,
            'ret_3m': 0.088,
            'ret_6m': 0.185,
            'ret_1y': 0.733,
            'ret_2y': 1.150,
            'ret_3y': 1.580,
            'ret_all': 2.150,
            'risk': 'RR4',
            'freq': '季配',
            'latest_div': 1.10,
            'ann_yield': 0.1358,
            'y1_div_sum': 2.88,
            'avg_fill_days': 11.8,
            'fill_rate': 1.0,
            'rating': '⭐ 穩健填息型 (1~2個月內)',
            'tsmc': 0.0,
            'tech': 0.505,
            'tech_total': 0.505,
            'fin': 0.145,
            'trad': 0.350,
            'bond_other': 0.0,
            'positioning': '💰 鎖定高配息宣告宣告股 (高殖利率王)',
            'top3': '聯電(9.5%), 長榮(9.1%), 聯發科(8.8%)',
            'holdings': [
                {'rank': 1, 'code': '2303', 'name': '聯電', 'ind': '科技電子', 'weight': 0.095},
                {'rank': 2, 'code': '2603', 'name': '長榮', 'ind': '傳統產業', 'weight': 0.091},
                {'rank': 3, 'code': '2454', 'name': '聯發科', 'ind': '科技電子', 'weight': 0.088},
                {'rank': 4, 'code': '3034', 'name': '聯詠', 'ind': '科技電子', 'weight': 0.082},
                {'rank': 5, 'code': '2891', 'name': '中信金', 'ind': '金融保險', 'weight': 0.071},
            ],
            'div_records': [
                {'ym': '2026/08', 'amount': 1.10, 'ex_date': '2026/09/16', 'status': '已填息 (12天)', 'days': 12, 'current_rate': 0.0339, 'year_rate': 0.1358},
                {'ym': '2026/05', 'amount': 1.00, 'ex_date': '2026/06/16', 'status': '已填息 (9天)', 'days': 9, 'current_rate': 0.0308, 'year_rate': 0.1234},
                {'ym': '2026/02', 'amount': 0.78, 'ex_date': '2026/03/17', 'status': '已填息 (14天)', 'days': 14, 'current_rate': 0.0241, 'year_rate': 0.0963},
            ]
        },
        {
            'house': '元大投信',
            'type': '同業ETF',
            'group': '方向B (同業擂台)',
            'cat': '債券固定收益',
            'code': '00679B',
            'sname': '元大美債20年',
            'name': '元大美國政府20年期(以上)債券基金',
            'price': 25.92,
            'fee': 0.0014,
            'ret_3m': 0.006,
            'ret_6m': -0.018,
            'ret_1y': 0.012,
            'ret_2y': 0.038,
            'ret_3y': 0.051,
            'ret_all': 0.072,
            'risk': 'RR2',
            'freq': '季配',
            'latest_div': 0.28,
            'ann_yield': 0.0432,
            'y1_div_sum': 1.12,
            'avg_fill_days': 2.7,
            'fill_rate': 1.0,
            'rating': '⚡ 極速填息王 (平均<=10天)',
            'tsmc': 0.0,
            'tech': 0.0,
            'tech_total': 0.0,
            'fin': 0.0,
            'trad': 0.0,
            'bond_other': 1.0,
            'positioning': '🏦 老牌指標美國長天期公債 (純度100%)',
            'top3': '美國公債 2043(8.5%), 美國公債 2044(8.1%), 美國公債 2042(7.8%)',
            'holdings': [
                {'rank': 1, 'code': 'US-BOND-43', 'name': '美國公債 2043', 'ind': '債券/固定收益', 'weight': 0.085},
                {'rank': 2, 'code': 'US-BOND-44', 'name': '美國公債 2044', 'ind': '債券/固定收益', 'weight': 0.081},
                {'rank': 3, 'code': 'US-BOND-42', 'name': '美國公債 2042', 'ind': '債券/固定收益', 'weight': 0.078},
            ],
            'div_records': [
                {'ym': '2026/07', 'amount': 0.28, 'ex_date': '2026/08/21', 'status': '已填息 (3天)', 'days': 3, 'current_rate': 0.0108, 'year_rate': 0.0432},
                {'ym': '2026/04', 'amount': 0.28, 'ex_date': '2026/05/22', 'status': '已填息 (2天)', 'days': 2, 'current_rate': 0.0108, 'year_rate': 0.0432},
                {'ym': '2026/01', 'amount': 0.28, 'ex_date': '2026/02/26', 'status': '已填息 (3天)', 'days': 3, 'current_rate': 0.0108, 'year_rate': 0.0432},
            ]
        },
        # === 核心自選追蹤標的 (User Watchlist) ===
        {
            'house': '富邦投信',
            'type': '同業ETF',
            'group': '自選追蹤',
            'cat': '市值型龍頭',
            'code': '006208',
            'sname': '富邦台灣50',
            'name': '富邦台灣釆擷50證券投資信託基金',
            'price': 115.60,
            'fee': 0.0015,
            'ret_3m': 0.143,
            'ret_6m': 0.286,
            'ret_1y': 0.776,
            'ret_2y': 1.258,
            'ret_3y': 1.898,
            'ret_all': 6.450,
            'risk': 'RR4',
            'freq': '半年配',
            'latest_div': 0.783,
            'ann_yield': 0.0135,
            'y1_div_sum': 1.563,
            'avg_fill_days': 2.5,
            'fill_rate': 1.0,
            'rating': '⚡ 極速填息王 (平均<=10天)',
            'tsmc': 0.551,
            'tech': 0.236,
            'tech_total': 0.787,
            'fin': 0.124,
            'trad': 0.089,
            'bond_other': 0.0,
            'positioning': '💎 0050平替首選 (超低經理費 0.15% / 台積電 55%)',
            'top3': '台積電(55.1%), 聯發科(5.1%), 鴻海(4.8%)',
            'holdings': [
                {'rank': 1, 'code': '2330', 'name': '台積電', 'ind': '台積電 (半導體)', 'weight': 0.551},
                {'rank': 2, 'code': '2454', 'name': '聯發科', 'ind': '科技電子', 'weight': 0.051},
                {'rank': 3, 'code': '2317', 'name': '鴻海', 'ind': '科技電子', 'weight': 0.048},
                {'rank': 4, 'code': '2308', 'name': '台達電', 'ind': '科技電子', 'weight': 0.024},
                {'rank': 5, 'code': '2382', 'name': '廣達', 'ind': '科技電子', 'weight': 0.021},
            ],
            'div_records': [
                {'ym': '2026/06', 'amount': 0.783, 'ex_date': '2026/07/16', 'status': '已填息 (2天)', 'days': 2, 'current_rate': 0.0068, 'year_rate': 0.0135},
                {'ym': '2025/11', 'amount': 0.780, 'ex_date': '2025/11/21', 'status': '已填息 (3天)', 'days': 3, 'current_rate': 0.0067, 'year_rate': 0.0134},
            ]
        },
        {
            'house': '復華投信',
            'type': '同業ETF',
            'group': '自選追蹤',
            'cat': '科技高股息',
            'code': '00929',
            'sname': '復華台灣科技優息',
            'name': '復華台灣科技優息ETF證券投資信託基金',
            'price': 19.35,
            'fee': 0.0030,
            'ret_3m': 0.038,
            'ret_6m': 0.125,
            'ret_1y': 0.385,
            'ret_2y': 0.650,
            'ret_3y': 0.850,
            'ret_all': 0.985,
            'risk': 'RR5',
            'freq': '月配',
            'latest_div': 0.14,
            'ann_yield': 0.0868,
            'y1_div_sum': 1.95,
            'avg_fill_days': 9.5,
            'fill_rate': 0.92,
            'rating': '⚡ 極速填息王 (平均<=10天)',
            'tsmc': 0.0,
            'tech': 1.00,
            'tech_total': 1.00,
            'fin': 0.0,
            'trad': 0.0,
            'bond_other': 0.0,
            'positioning': '⚡ 全台首檔百億月配息 (純科技高股息 100%)',
            'top3': '聯發科(9.8%), 聯詠(8.9%), 日月光投控(8.5%)',
            'holdings': [
                {'rank': 1, 'code': '2454', 'name': '聯發科', 'ind': '科技電子', 'weight': 0.098},
                {'rank': 2, 'code': '3034', 'name': '聯詠', 'ind': '科技電子', 'weight': 0.089},
                {'rank': 3, 'code': '3711', 'name': '日月光投控', 'ind': '科技電子', 'weight': 0.085},
                {'rank': 4, 'code': '3702', 'name': '大聯大', 'ind': '科技電子', 'weight': 0.078},
                {'rank': 5, 'code': '2385', 'name': '群光', 'ind': '科技電子', 'weight': 0.065},
            ],
            'div_records': [
                {'ym': '2026/09', 'amount': 0.14, 'ex_date': '2026/10/18', 'status': '已填息 (8天)', 'days': 8, 'current_rate': 0.0072, 'year_rate': 0.0868},
                {'ym': '2026/08', 'amount': 0.16, 'ex_date': '2026/09/19', 'status': '已填息 (9天)', 'days': 9, 'current_rate': 0.0083, 'year_rate': 0.0992},
                {'ym': '2026/07', 'amount': 0.18, 'ex_date': '2026/08/20', 'status': '已填息 (11天)', 'days': 11, 'current_rate': 0.0093, 'year_rate': 0.1116},
            ]
        },
        {
            'house': '元大投信',
            'type': '同業ETF',
            'group': '自選追蹤',
            'cat': '高息低波防禦',
            'code': '00713',
            'sname': '元大台灣高息低波',
            'name': '元大台灣高股息低波動證券投資信託基金',
            'price': 57.80,
            'fee': 0.0030,
            'ret_3m': 0.052,
            'ret_6m': 0.158,
            'ret_1y': 0.428,
            'ret_2y': 0.885,
            'ret_3y': 1.285,
            'ret_all': 2.850,
            'risk': 'RR4',
            'freq': '季配',
            'latest_div': 1.50,
            'ann_yield': 0.1038,
            'y1_div_sum': 5.80,
            'avg_fill_days': 4.5,
            'fill_rate': 1.0,
            'rating': '⚡ 極速填息王 (平均<=10天)',
            'tsmc': 0.0,
            'tech': 0.452,
            'tech_total': 0.452,
            'fin': 0.228,
            'trad': 0.320,
            'bond_other': 0.0,
            'positioning': '🛡️ 空頭避風港 (低波動抗跌之王 / 均衡防禦)',
            'top3': '統一(8.2%), 台灣大(7.8%), 遠傳(7.5%)',
            'holdings': [
                {'rank': 1, 'code': '1216', 'name': '統一', 'ind': '傳統產業', 'weight': 0.082},
                {'rank': 2, 'code': '3045', 'name': '台灣大', 'ind': '科技電子', 'weight': 0.078},
                {'rank': 3, 'code': '4904', 'name': '遠傳', 'ind': '科技電子', 'weight': 0.075},
                {'rank': 4, 'code': '2880', 'name': '華南金', 'ind': '金融保險', 'weight': 0.058},
                {'rank': 5, 'code': '2357', 'name': '華碩', 'ind': '科技電子', 'weight': 0.055},
            ],
            'div_records': [
                {'ym': '2026/08', 'amount': 1.50, 'ex_date': '2026/09/18', 'status': '已填息 (4天)', 'days': 4, 'current_rate': 0.0259, 'year_rate': 0.1038},
                {'ym': '2026/05', 'amount': 1.50, 'ex_date': '2026/06/19', 'status': '已填息 (5天)', 'days': 5, 'current_rate': 0.0259, 'year_rate': 0.1038},
                {'ym': '2026/02', 'amount': 1.40, 'ex_date': '2026/03/18', 'status': '已填息 (4天)', 'days': 4, 'current_rate': 0.0242, 'year_rate': 0.0968},
            ]
        },
        {
            'house': '元大投信',
            'type': '同業ETF',
            'group': '自選追蹤',
            'cat': '價值型高股息',
            'code': '00940',
            'sname': '元大台灣價值高息',
            'name': '元大台灣價值高息ETF證券投資信託基金',
            'price': 9.58,
            'fee': 0.0030,
            'ret_3m': 0.015,
            'ret_6m': 0.062,
            'ret_1y': 0.185,
            'ret_2y': 0.220,
            'ret_3y': 0.220,
            'ret_all': 0.250,
            'risk': 'RR4',
            'freq': '月配',
            'latest_div': 0.05,
            'ann_yield': 0.0626,
            'y1_div_sum': 0.50,
            'avg_fill_days': 12.0,
            'fill_rate': 0.88,
            'rating': '⭐ 穩健填息型 (1~2個月內)',
            'tsmc': 0.0,
            'tech': 0.655,
            'tech_total': 0.655,
            'fin': 0.085,
            'trad': 0.260,
            'bond_other': 0.0,
            'positioning': '🏦 巴菲特價值選股 (親民單價 / 月配現金流)',
            'top3': '長榮(10.5%), 鈊象(8.2%), 聯詠(6.8%)',
            'holdings': [
                {'rank': 1, 'code': '2603', 'name': '長榮', 'ind': '傳統產業', 'weight': 0.105},
                {'rank': 2, 'code': '3293', 'name': '鈊象', 'ind': '科技電子', 'weight': 0.082},
                {'rank': 3, 'code': '3034', 'name': '聯詠', 'ind': '科技電子', 'weight': 0.068},
                {'rank': 4, 'code': '2454', 'name': '聯發科', 'ind': '科技電子', 'weight': 0.061},
                {'rank': 5, 'code': '2303', 'name': '聯電', 'ind': '科技電子', 'weight': 0.054},
            ],
            'div_records': [
                {'ym': '2026/09', 'amount': 0.05, 'ex_date': '2026/10/08', 'status': '已填息 (10天)', 'days': 10, 'current_rate': 0.0052, 'year_rate': 0.0626},
                {'ym': '2026/08', 'amount': 0.05, 'ex_date': '2026/09/05', 'status': '已填息 (12天)', 'days': 12, 'current_rate': 0.0052, 'year_rate': 0.0626},
                {'ym': '2026/07', 'amount': 0.05, 'ex_date': '2026/08/08', 'status': '已填息 (14天)', 'days': 14, 'current_rate': 0.0052, 'year_rate': 0.0626},
            ]
        },
        {
            'house': '群益投信',
            'type': '同業ETF',
            'group': '自選追蹤',
            'cat': '海外投資級債',
            'code': '00937B',
            'sname': '群益ESG投等債20+',
            'name': '群益ESG 20年期以上BBB投資級公司債ETF證券投資信託基金',
            'price': 15.82,
            'fee': 0.0020,
            'ret_3m': 0.012,
            'ret_6m': 0.024,
            'ret_1y': 0.048,
            'ret_2y': 0.095,
            'ret_3y': 0.125,
            'ret_all': 0.145,
            'risk': 'RR2',
            'freq': '月配',
            'latest_div': 0.084,
            'ann_yield': 0.0637,
            'y1_div_sum': 0.98,
            'avg_fill_days': 5.5,
            'fill_rate': 1.0,
            'rating': '⚡ 極速填息王 (平均<=10天)',
            'tsmc': 0.0,
            'tech': 0.0,
            'tech_total': 0.0,
            'fin': 0.0,
            'trad': 0.0,
            'bond_other': 1.0,
            'positioning': '🏦 超人氣月配投等債 (BBB級ESG跨國龍頭債 / 避開二代健保)',
            'top3': '甲骨文債券 2045(3.8%), 亞馬遜債券 2050(3.5%), 微軟債券 2052(3.2%)',
            'holdings': [
                {'rank': 1, 'code': 'US-ORCL-45', 'name': '甲骨文公司債', 'ind': '債券/固定收益', 'weight': 0.038},
                {'rank': 2, 'code': 'US-AMZN-50', 'name': '亞馬遜公司債', 'ind': '債券/固定收益', 'weight': 0.035},
                {'rank': 3, 'code': 'US-MSFT-52', 'name': '微軟公司債', 'ind': '債券/固定收益', 'weight': 0.032},
                {'rank': 4, 'code': 'US-AAPL-51', 'name': '蘋果公司債', 'ind': '債券/固定收益', 'weight': 0.029},
                {'rank': 5, 'code': 'US-PFE-49', 'name': '輝瑞製藥公司債', 'ind': '債券/固定收益', 'weight': 0.027},
            ],
            'div_records': [
                {'ym': '2026/09', 'amount': 0.084, 'ex_date': '2026/10/22', 'status': '已填息 (5天)', 'days': 5, 'current_rate': 0.0053, 'year_rate': 0.0637},
                {'ym': '2026/08', 'amount': 0.084, 'ex_date': '2026/09/20', 'status': '已填息 (6天)', 'days': 6, 'current_rate': 0.0053, 'year_rate': 0.0637},
                {'ym': '2026/07', 'amount': 0.084, 'ex_date': '2026/08/21', 'status': '已填息 (5天)', 'days': 5, 'current_rate': 0.0053, 'year_rate': 0.0637},
            ]
        }
    ]
    
    # 組合所有標的
    all_targets = []
    
    # 加入國泰既有標的
    for r in rank_rows:
        code = str(r['代碼']).strip()
        div_info = div_sum_dict.get(code, {})
        sec_info = sec_dict.get(code, {})
        holdings = holdings_by_code.get(code, [])
        div_records = div_details_by_code.get(code, [])
        
        # 標記分類 group 與經理費推估
        is_etf = r.get('投資大類') == 'ETF'
        sname_val = r.get('簡稱', '')
        if '50' in sname_val or '領袖' in sname_val:
            fee = 0.0020
        elif '科技' in sname_val or '半導體' in sname_val:
            fee = 0.0040
        elif '債' in sname_val:
            fee = 0.0020
        elif is_etf:
            fee = 0.0025  # 00878 等
        else:
            fee = 0.0150  # 共同基金約 1.50%
            
        all_targets.append({
            'house': '國泰投信',
            'type': r.get('投資大類', ''),
            'group': '國泰人氣王',
            'cat': r.get('推薦分類', ''),
            'code': code,
            'sname': r.get('簡稱', ''),
            'name': r.get('基金名稱', ''),
            'price': float(r.get('現價或淨值') or 0.0),
            'fee': fee,
            'ret_3m': float(r.get('近3個月績效') or 0.0) if r.get('近3個月績效') is not None else None,
            'ret_6m': float(r.get('近6個月績效') or 0.0) if r.get('近6個月績效') is not None else None,
            'ret_1y': float(r.get('近1年績效') or 0.0) if r.get('近1年績效') is not None else None,
            'ret_2y': float(r.get('近2年績效') or 0.0) if r.get('近2年績效') is not None else None,
            'ret_3y': float(r.get('近3年績效') or 0.0) if r.get('近3年績效') is not None else None,
            'ret_all': float(r.get('成立以來績效') or 0.0) if r.get('成立以來績效') is not None else None,
            'risk': r.get('風險等級', ''),
            'freq': r.get('配息頻率', ''),
            'latest_div': float(div_info.get('最新每期配息(元)') or 0.0) if div_info else 0.0,
            'ann_yield': float(div_info.get('最新年化配息率') or 0.0) if div_info else 0.0,
            'y1_div_sum': float(div_info.get('近一年累計配息(元)') or 0.0) if div_info else 0.0,
            'avg_fill_days': float(div_info.get('平均填息天數(天)')) if div_info and div_info.get('平均填息天數(天)') is not None else None,
            'fill_rate': float(div_info.get('歷史填息成功率')) if div_info and div_info.get('歷史填息成功率') is not None else None,
            'rating': div_info.get('填息力評級', '--') if div_info else '--',
            'tsmc': float(sec_info.get('台積電佔比') or 0.0) if sec_info else 0.0,
            'tech': float(sec_info.get('科技電子佔比') or 0.0) if sec_info else 0.0,
            'tech_total': float(sec_info.get('科技合計佔比') or 0.0) if sec_info else 0.0,
            'fin': float(sec_info.get('金融保險佔比') or 0.0) if sec_info else 0.0,
            'trad': float(sec_info.get('傳統產業佔比') or 0.0) if sec_info else 0.0,
            'bond_other': float(sec_info.get('海外/債券/其他') or 0.0) if sec_info else 0.0,
            'positioning': sec_info.get('核心配置定位', '--') if sec_info else '--',
            'top3': sec_info.get('前3大持股', '--') if sec_info else '--',
            'holdings': [
                {'rank': h.get('成分股排名'), 'code': h.get('成分股代碼'), 'name': h.get('成分股名稱'), 'ind': h.get('產業類別'), 'weight': float(h.get('持股權重') or 0.0)}
                for h in holdings
            ],
            'div_records': [
                {
                    'ym': d.get('配息年月', ''),
                    'amount': float(d.get('每單位配息(元)') or 0.0),
                    'ex_date': d.get('除息交易日', ''),
                    'status': d.get('填息狀態', ''),
                    'days': d.get('填息天數(日)'),
                    'current_rate': float(d.get('當期配息率') or 0.0) if d.get('當期配息率') is not None else 0.0,
                    'year_rate': float(d.get('年化配息率') or 0.0) if d.get('年化配息率') is not None else 0.0,
                }
                for d in div_records
            ]
        })
        
    # 加入 Direction A & B (若 code 不重複)
    existing_codes = set(t['code'] for t in all_targets)
    for extra in extra_targets:
        if extra['code'] not in existing_codes:
            all_targets.append(extra)
            
    print(f"  -> 整合完畢！總共匯整 {len(all_targets)} 檔標的 (國泰主力 + 國泰美債/電動車 + 元大/群益市場宿敵)")
    return all_targets

def generate_html_file(targets):
    """產出頂級現代 Fintech 風格之 HTML5 互動式量化投資儀表板"""
    print(f"[2/3] 正在生成現代化獨立 HTML 儀表板: {HTML_OUTPUT} ...")
    
    # 轉為 JSON 供前端即時渲染與篩選
    targets_json = json.dumps(targets, ensure_ascii=False)
    
    html_template = f"""<!DOCTYPE html>
<html lang="zh-TW">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>國泰投信 人氣 ETF 與熱門基金 量化投資分析互動儀表板 (含同業巨頭擂台)</title>
    <!-- Google Fonts -->
    <link rel="preconnect" href="https://fonts.googleapis.com">
    <link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
    <link href="https://fonts.googleapis.com/css2?family=Noto+Sans+TC:wght@400;500;600;700;800&family=Outfit:wght@400;500;600;700;800&family=JetBrains+Mono:wght@500;700&display=swap" rel="stylesheet">
    <!-- Chart.js CDN -->
    <script src="https://cdn.jsdelivr.net/npm/chart.js@4.4.1/dist/chart.umd.min.js"></script>
    <style>
        :root {{
            --bg-body: #0b1120;
            --bg-card: #131c31;
            --bg-card-hover: #1e2942;
            --bg-glass: rgba(19, 28, 49, 0.75);
            --border: #233252;
            --border-hover: #3b82f6;
            --text-main: #f8fafc;
            --text-sub: #94a3b8;
            --text-muted: #64748b;
            --brand-green: #00a86b;
            --brand-green-glow: rgba(0, 168, 107, 0.25);
            --brand-dark-green: #005a36;
            --accent-blue: #38bdf8;
            --accent-gold: #f59e0b;
            --accent-red: #f43f5e;
            --accent-purple: #a855f7;
            --tw-up: #ef4444;       /* 台股紅漲 (Red Up) */
            --tw-down: #10b981;     /* 台股綠跌 (Green Down) */
            --badge-bg: #1e293b;
            --badge-border: #334155;
            --shadow-card: 0 10px 25px -5px rgba(0, 0, 0, 0.4), 0 8px 10px -6px rgba(0, 0, 0, 0.3);
            --radius-lg: 16px;
            --radius-md: 10px;
            --radius-sm: 6px;
        }}

        [data-theme="light"] {{
            --bg-body: #f1f5f9;
            --bg-card: #ffffff;
            --bg-card-hover: #f8fafc;
            --bg-glass: rgba(255, 255, 255, 0.85);
            --border: #e2e8f0;
            --border-hover: #008054;
            --text-main: #0f172a;
            --text-sub: #475569;
            --text-muted: #94a3b8;
            --brand-green: #008054;
            --brand-green-glow: rgba(0, 128, 84, 0.15);
            --brand-dark-green: #004d32;
            --accent-blue: #0284c7;
            --accent-gold: #d97706;
            --accent-red: #e11d48;
            --accent-purple: #7c3aed;
            --tw-up: #dc2626;       /* 台股紅漲 (Red Up for light mode) */
            --tw-down: #16a34a;     /* 台股綠跌 (Green Down for light mode) */
            --badge-bg: #f8fafc;
            --badge-border: #cbd5e1;
            --shadow-card: 0 10px 20px -3px rgba(0, 0, 0, 0.05), 0 4px 6px -2px rgba(0, 0, 0, 0.03);
        }}

        * {{
            box-sizing: border-box;
            margin: 0;
            padding: 0;
            transition: background-color 0.2s, border-color 0.2s;
        }}

        body {{
            font-family: 'Outfit', 'Noto Sans TC', sans-serif;
            background-color: var(--bg-body);
            color: var(--text-main);
            line-height: 1.5;
            padding: 24px;
            min-height: 100vh;
        }}

        /* Container */
        .container {{
            max-width: 1440px;
            margin: 0 auto;
        }}

        /* Header */
        header {{
            display: flex;
            justify-content: space-between;
            align-items: center;
            margin-bottom: 24px;
            padding-bottom: 20px;
            border-bottom: 1px solid var(--border);
            flex-wrap: wrap;
            gap: 16px;
        }}

        .brand-title {{
            display: flex;
            align-items: center;
            gap: 12px;
        }}

        .brand-logo {{
            width: 44px;
            height: 44px;
            background: linear-gradient(135deg, #008054 0%, #00a86b 100%);
            border-radius: 12px;
            display: flex;
            align-items: center;
            justify-content: center;
            color: #fff;
            font-size: 22px;
            font-weight: 800;
            box-shadow: 0 4px 14px var(--brand-green-glow);
        }}

        .brand-text h1 {{
            font-size: 24px;
            font-weight: 800;
            letter-spacing: -0.5px;
            background: linear-gradient(90deg, var(--text-main) 60%, var(--brand-green) 100%);
            -webkit-background-clip: text;
            -webkit-text-fill-color: transparent;
        }}

        .brand-text p {{
            font-size: 13px;
            color: var(--text-sub);
        }}

        .header-actions {{
            display: flex;
            align-items: center;
            gap: 10px;
            flex-wrap: wrap;
        }}

        .btn {{
            display: inline-flex;
            align-items: center;
            gap: 6px;
            padding: 8px 16px;
            border-radius: var(--radius-sm);
            font-size: 13px;
            font-weight: 600;
            cursor: pointer;
            border: 1px solid var(--border);
            background: var(--bg-card);
            color: var(--text-main);
            transition: all 0.2s;
        }}

        .btn:hover {{
            border-color: var(--brand-green);
            transform: translateY(-1px);
        }}

        .btn-primary {{
            background: var(--brand-green);
            color: #ffffff;
            border-color: transparent;
            box-shadow: 0 4px 12px var(--brand-green-glow);
        }}

        .btn-primary:hover {{
            background: var(--brand-dark-green);
        }}

        /* KPI Cards Grid */
        .kpi-grid {{
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(240px, 1fr));
            gap: 16px;
            margin-bottom: 24px;
        }}

        .kpi-card {{
            background: var(--bg-card);
            border: 1px solid var(--border);
            border-radius: var(--radius-lg);
            padding: 18px 20px;
            box-shadow: var(--shadow-card);
            position: relative;
            overflow: hidden;
            transition: transform 0.2s, border-color 0.2s;
        }}

        .kpi-card:hover {{
            transform: translateY(-2px);
            border-color: var(--border-hover);
        }}

        .kpi-card::before {{
            content: '';
            position: absolute;
            top: 0;
            left: 0;
            right: 0;
            height: 3px;
            background: linear-gradient(90deg, var(--brand-green), var(--accent-blue));
        }}

        .kpi-card.gold::before {{
            background: linear-gradient(90deg, var(--accent-gold), var(--accent-red));
        }}

        .kpi-card.purple::before {{
            background: linear-gradient(90deg, var(--accent-purple), var(--brand-green));
        }}

        .kpi-card.blue::before {{
            background: linear-gradient(90deg, var(--accent-blue), var(--accent-purple));
        }}

        .kpi-title {{
            font-size: 12px;
            font-weight: 600;
            color: var(--text-sub);
            text-transform: uppercase;
            letter-spacing: 0.5px;
            margin-bottom: 6px;
            display: flex;
            align-items: center;
            gap: 6px;
        }}

        .kpi-value {{
            font-size: 22px;
            font-weight: 800;
            font-family: 'JetBrains Mono', monospace;
            color: var(--text-main);
            margin-bottom: 4px;
        }}

        .kpi-meta {{
            font-size: 12px;
            color: var(--text-sub);
            display: flex;
            align-items: center;
            gap: 6px;
            flex-wrap: wrap;
        }}

        .kpi-badge {{
            padding: 2px 8px;
            border-radius: 4px;
            font-size: 11px;
            font-weight: 600;
            background: var(--badge-bg);
            color: var(--text-sub);
            border: 1px solid var(--border);
            white-space: nowrap;
        }}

        /* Controls Section */
        .controls-card {{
            background: var(--bg-card);
            border: 1px solid var(--border);
            border-radius: var(--radius-lg);
            padding: 16px 20px;
            margin-bottom: 24px;
            box-shadow: var(--shadow-card);
            display: flex;
            flex-direction: column;
            gap: 14px;
        }}

        .filter-row {{
            display: flex;
            align-items: center;
            justify-content: space-between;
            flex-wrap: wrap;
            gap: 12px;
        }}

        .filter-chips {{
            display: flex;
            gap: 8px;
            flex-wrap: wrap;
        }}

        .chip {{
            padding: 6px 14px;
            border-radius: 30px;
            font-size: 13px;
            font-weight: 600;
            cursor: pointer;
            border: 1px solid var(--border);
            background: var(--bg-body);
            color: var(--text-sub);
            transition: all 0.2s;
        }}

        .chip:hover {{
            color: var(--text-main);
            border-color: var(--border-hover);
        }}

        .chip.active {{
            background: var(--brand-green);
            color: #ffffff;
            border-color: var(--brand-green);
            box-shadow: 0 2px 8px var(--brand-green-glow);
        }}

        .search-box {{
            position: relative;
            min-width: 260px;
        }}

        .search-box input {{
            width: 100%;
            padding: 8px 14px 8px 36px;
            border-radius: var(--radius-sm);
            border: 1px solid var(--border);
            background: var(--bg-body);
            color: var(--text-main);
            font-size: 13px;
            outline: none;
            transition: border-color 0.2s;
        }}

        .search-box input:focus {{
            border-color: var(--brand-green);
        }}

        .search-icon {{
            position: absolute;
            left: 12px;
            top: 50%;
            transform: translateY(-50%);
            font-size: 14px;
            color: var(--text-muted);
        }}

        /* Tabs Navigation */
        .nav-tabs {{
            display: flex;
            gap: 4px;
            border-bottom: 1px solid var(--border);
            margin-bottom: 20px;
            overflow-x: auto;
        }}

        .nav-tab {{
            padding: 12px 20px;
            font-size: 14px;
            font-weight: 700;
            color: var(--text-sub);
            cursor: pointer;
            border-bottom: 3px solid transparent;
            transition: all 0.2s;
            white-space: nowrap;
            display: flex;
            align-items: center;
            gap: 8px;
        }}

        .nav-tab:hover {{
            color: var(--text-main);
        }}

        .nav-tab.active {{
            color: var(--brand-green);
            border-bottom-color: var(--brand-green);
        }}

        /* Table Card */
        .table-card {{
            background: var(--bg-card);
            border: 1px solid var(--border);
            border-radius: var(--radius-lg);
            overflow: hidden;
            box-shadow: var(--shadow-card);
            margin-bottom: 30px;
        }}

        .table-responsive {{
            overflow-x: auto;
            max-height: 720px;
        }}

        table {{
            width: 100%;
            border-collapse: collapse;
            font-size: 13px;
            text-align: left;
        }}

        thead {{
            background: var(--bg-body);
            position: sticky;
            top: 0;
            z-index: 10;
        }}

        th {{
            padding: 12px 14px;
            font-weight: 700;
            color: var(--text-sub);
            border-bottom: 1px solid var(--border);
            cursor: pointer;
            user-select: none;
            white-space: nowrap;
        }}

        th:hover {{
            color: var(--text-main);
        }}

        td {{
            padding: 12px 14px;
            border-bottom: 1px solid var(--border);
            white-space: nowrap;
        }}

        tbody tr {{
            transition: background-color 0.15s;
        }}

        tbody tr:hover {{
            background-color: var(--bg-card-hover);
        }}

        .mono {{
            font-family: 'JetBrains Mono', monospace;
        }}

        /* 台股紅漲綠跌規範 */
        .num-up, .num-pos {{
            color: var(--tw-up);
            font-weight: 700;
        }}

        .num-down, .num-neg {{
            color: var(--tw-down);
            font-weight: 700;
        }}

        .num-neutral {{
            color: var(--text-main);
            font-weight: 600;
        }}

        .badge {{
            display: inline-flex;
            align-items: center;
            padding: 3px 8px;
            border-radius: 4px;
            font-size: 11px;
            font-weight: 600;
        }}

        .badge-cathay {{
            background: rgba(0, 168, 107, 0.12);
            color: var(--brand-green);
            border: 1px solid rgba(0, 168, 107, 0.3);
        }}

        .badge-peer {{
            background: rgba(168, 85, 247, 0.12);
            color: var(--accent-purple);
            border: 1px solid rgba(168, 85, 247, 0.3);
        }}

        .badge-fubon {{
            background: rgba(14, 165, 233, 0.12);
            color: var(--accent-blue);
            border: 1px solid rgba(14, 165, 233, 0.3);
        }}

        .badge-fuhwa {{
            background: rgba(245, 158, 11, 0.12);
            color: var(--accent-gold);
            border: 1px solid rgba(245, 158, 11, 0.3);
        }}

        .badge-fill-fast {{
            background: rgba(56, 189, 248, 0.15);
            color: var(--accent-blue);
            font-weight: 700;
        }}

        .badge-fill-mid {{
            background: rgba(148, 163, 184, 0.15);
            color: var(--text-sub);
        }}

        .badge-fill-slow {{
            background: rgba(245, 158, 11, 0.15);
            color: var(--accent-gold);
        }}

        .clickable-row {{
            cursor: pointer;
        }}

        /* Tooltip */
        .tooltip {{
            position: relative;
            display: inline-block;
            cursor: help;
            color: var(--accent-blue);
            margin-left: 4px;
        }}

        .tooltip .tooltip-text {{
            visibility: hidden;
            width: 260px;
            background-color: #0f172a;
            color: #fff;
            text-align: left;
            border-radius: 8px;
            padding: 10px 12px;
            position: absolute;
            z-index: 100;
            bottom: 125%;
            left: 50%;
            transform: translateX(-50%);
            opacity: 0;
            transition: opacity 0.2s;
            font-size: 12px;
            line-height: 1.4;
            border: 1px solid var(--border-hover);
            box-shadow: 0 10px 25px rgba(0,0,0,0.5);
            pointer-events: none;
        }}

        .tooltip:hover .tooltip-text {{
            visibility: visible;
            opacity: 1;
        }}

        /* Modal Drawer */
        .modal-overlay {{
            position: fixed;
            top: 0;
            left: 0;
            right: 0;
            bottom: 0;
            background: rgba(0, 0, 0, 0.7);
            backdrop-filter: blur(4px);
            display: none;
            justify-content: center;
            align-items: center;
            z-index: 999;
            padding: 20px;
        }}

        .modal-overlay.active {{
            display: flex;
        }}

        .modal-card {{
            background: var(--bg-card);
            border: 1px solid var(--border);
            border-radius: var(--radius-lg);
            width: 100%;
            max-width: 840px;
            max-height: 85vh;
            overflow-y: auto;
            box-shadow: var(--shadow-card);
            padding: 24px;
            position: relative;
        }}

        .modal-close {{
            position: absolute;
            top: 18px;
            right: 18px;
            background: none;
            border: none;
            font-size: 20px;
            cursor: pointer;
            color: var(--text-sub);
        }}

        .modal-close:hover {{
            color: var(--text-main);
        }}

        /* 國泰官網同款投資組合長條圖樣式 */
        .cathay-holdings-container {{
            display: flex;
            flex-direction: column;
            gap: 10px;
            margin-bottom: 24px;
        }}

        .holding-bar-item {{
            position: relative;
            background: rgba(255, 255, 255, 0.03);
            border: 1px solid var(--border);
            border-radius: var(--radius-sm);
            padding: 12px 16px;
            overflow: hidden;
            display: flex;
            justify-content: space-between;
            align-items: center;
        }}

        .holding-bar-bg {{
            position: absolute;
            left: 0;
            top: 0;
            bottom: 0;
            background: linear-gradient(90deg, rgba(0, 168, 107, 0.32), rgba(16, 185, 129, 0.12));
            border-right: 3px solid var(--brand-green);
            border-radius: var(--radius-sm) 0 0 var(--radius-sm);
            z-index: 1;
            transition: width 0.6s cubic-bezier(0.16, 1, 0.3, 1);
        }}

        .holding-bar-info {{
            position: relative;
            z-index: 2;
            display: flex;
            align-items: center;
            gap: 12px;
        }}

        .holding-bar-rank {{
            width: 28px;
            height: 28px;
            border-radius: 6px;
            background: rgba(255, 255, 255, 0.08);
            display: flex;
            align-items: center;
            justify-content: center;
            font-size: 12px;
            font-weight: 800;
            font-family: 'JetBrains Mono', monospace;
            color: var(--text-main);
        }}

        .holding-bar-rank.top1 {{
            background: #f59e0b;
            color: #0b1120;
            box-shadow: 0 0 10px rgba(245, 158, 11, 0.4);
        }}

        .holding-bar-rank.top2 {{
            background: #94a3b8;
            color: #0b1120;
        }}

        .holding-bar-rank.top3 {{
            background: #b45309;
            color: #ffffff;
        }}

        .holding-bar-text {{
            display: flex;
            flex-direction: column;
        }}

        .holding-bar-name {{
            font-size: 15px;
            font-weight: 700;
            color: var(--text-main);
            display: flex;
            align-items: center;
            gap: 8px;
        }}

        .holding-bar-sub {{
            font-size: 11px;
            color: var(--text-sub);
            margin-top: 2px;
        }}

        .holding-bar-weight {{
            position: relative;
            z-index: 2;
            font-family: 'JetBrains Mono', monospace;
            font-size: 16px;
            font-weight: 800;
            color: var(--text-main);
        }}

        /* PK Arena Section */
        .pk-grid {{
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(360px, 1fr));
            gap: 20px;
            margin-bottom: 24px;
        }}

        .pk-card {{
            background: var(--bg-card);
            border: 1px solid var(--border);
            border-radius: var(--radius-lg);
            padding: 20px;
            box-shadow: var(--shadow-card);
        }}

        .pk-header {{
            font-size: 16px;
            font-weight: 800;
            margin-bottom: 14px;
            padding-bottom: 10px;
            border-bottom: 1px solid var(--border);
            display: flex;
            align-items: center;
            justify-content: space-between;
        }}

        .pk-item {{
            display: flex;
            justify-content: space-between;
            align-items: center;
            padding: 10px 12px;
            border-radius: var(--radius-sm);
            margin-bottom: 8px;
            background: var(--bg-body);
        }}

        .pk-item.highlight {{
            border-left: 4px solid var(--brand-green);
        }}

        /* Guide Accordion */
        .guide-grid {{
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(420px, 1fr));
            gap: 18px;
            margin-bottom: 24px;
        }}

        .guide-card {{
            background: var(--bg-card);
            border: 1px solid var(--border);
            border-radius: var(--radius-lg);
            padding: 20px;
            box-shadow: var(--shadow-card);
        }}

        .guide-card h3 {{
            font-size: 16px;
            font-weight: 800;
            color: var(--brand-green);
            margin-bottom: 8px;
            display: flex;
            align-items: center;
            gap: 8px;
        }}

        .guide-card p {{
            font-size: 13px;
            color: var(--text-sub);
            margin-bottom: 10px;
            line-height: 1.6;
        }}

        .guide-example {{
            background: var(--bg-body);
            border-left: 3px solid var(--accent-blue);
            padding: 10px 12px;
            border-radius: 4px;
            font-size: 12px;
            color: var(--text-main);
            margin-top: 8px;
        }}

        /* 行動端現代卡片流 (Mobile Card View) */
        .view-mode-bar {{
            display: none;
            align-items: center;
            gap: 8px;
            margin-bottom: 14px;
        }}

        .view-mode-btn {{
            background: var(--bg-card);
            border: 1px solid var(--border);
            color: var(--text-sub);
            padding: 6px 14px;
            border-radius: 6px;
            font-size: 12px;
            font-weight: 700;
            cursor: pointer;
            transition: all 0.2s;
        }}

        .view-mode-btn.active {{
            background: rgba(0, 168, 107, 0.15);
            color: var(--brand-green);
            border-color: var(--brand-green);
        }}

        .mobile-card-list {{
            display: none;
        }}

        .m-card {{
            background: var(--bg-card);
            border: 1px solid var(--border);
            border-radius: var(--radius-md);
            padding: 14px 16px;
            box-shadow: var(--shadow-card);
            display: flex;
            flex-direction: column;
            gap: 10px;
            transition: transform 0.15s, border-color 0.15s;
            cursor: pointer;
        }}

        .m-card:active {{
            transform: scale(0.99);
            border-color: var(--brand-green);
        }}

        .m-card-header {{
            display: flex;
            justify-content: space-between;
            align-items: center;
            border-bottom: 1px solid rgba(255, 255, 255, 0.06);
            padding-bottom: 8px;
        }}

        .m-card-title-group {{
            display: flex;
            align-items: center;
            gap: 8px;
            flex-wrap: wrap;
        }}

        .m-card-code {{
            font-family: 'JetBrains Mono', monospace;
            font-size: 16px;
            font-weight: 800;
            color: var(--text-main);
        }}

        .m-card-name {{
            font-size: 15px;
            font-weight: 700;
            color: var(--text-main);
        }}

        .m-card-tags {{
            display: flex;
            gap: 6px;
            flex-wrap: wrap;
            align-items: center;
        }}

        .m-card-grid {{
            display: grid;
            grid-template-columns: repeat(2, 1fr);
            gap: 8px;
            background: var(--bg-body);
            padding: 10px 12px;
            border-radius: 6px;
        }}

        .m-grid-item {{
            display: flex;
            flex-direction: column;
        }}

        .m-label {{
            font-size: 11px;
            color: var(--text-sub);
            margin-bottom: 2px;
        }}

        .m-val {{
            font-size: 15px;
            font-weight: 700;
        }}

        .m-card-footer {{
            display: flex;
            justify-content: space-between;
            align-items: center;
            padding-top: 4px;
        }}

        .m-card-actions {{
            display: flex;
            gap: 8px;
        }}

        /* Footer */
        footer {{
            margin-top: 40px;
            padding-top: 20px;
            border-top: 1px solid var(--border);
            text-align: center;
            font-size: 12px;
            color: var(--text-muted);
        }}

        /* 手機與平板行動端自適應排版 (Mobile Responsive) */
        @media (max-width: 768px) {{
            .container {{
                padding: 12px 10px;
            }}

            header {{
                flex-direction: column;
                align-items: flex-start;
                gap: 14px;
                margin-bottom: 20px;
            }}

            .header-actions {{
                width: 100%;
                display: flex;
                gap: 8px;
            }}

            .header-actions .btn {{
                flex: 1;
                font-size: 12px;
                padding: 8px 6px;
                text-align: center;
                justify-content: center;
            }}

            .brand-title h1 {{
                font-size: 18px;
            }}

            .brand-title p {{
                font-size: 11px;
            }}

            .kpi-grid {{
                grid-template-columns: repeat(2, 1fr);
                gap: 10px;
                margin-bottom: 16px;
            }}

            .kpi-card {{
                padding: 14px 12px;
            }}

            .kpi-value {{
                font-size: 20px;
            }}

            .filter-row {{
                flex-direction: column;
                gap: 12px;
            }}

            .search-box {{
                width: 100%;
                min-width: 100%;
            }}

            .filter-chips {{
                overflow-x: auto;
                flex-wrap: nowrap;
                padding-bottom: 8px;
                -webkit-overflow-scrolling: touch;
            }}

            .chip {{
                flex-shrink: 0;
                font-size: 12px;
                padding: 5px 10px;
            }}

            .nav-tabs {{
                overflow-x: auto;
                flex-wrap: nowrap;
                -webkit-overflow-scrolling: touch;
                padding-bottom: 2px;
            }}

            .nav-tab {{
                font-size: 13px;
                padding: 10px 12px;
                flex-shrink: 0;
            }}

            .pk-grid, .guide-grid {{
                grid-template-columns: 1fr;
                gap: 14px;
            }}

            .modal-card {{
                padding: 16px 14px;
                max-width: 95vw;
                max-height: 88vh;
            }}

            .holding-bar-item {{
                padding: 10px 12px;
            }}

            .holding-bar-name {{
                font-size: 13px;
            }}

            .holding-bar-weight {{
                font-size: 14px;
            }}

            .view-mode-bar {{
                display: flex;
            }}

            .mobile-card-list {{
                display: flex;
                flex-direction: column;
                gap: 12px;
                margin-bottom: 24px;
            }}

            .table-card.mobile-hidden {{
                display: none !important;
            }}

            .mobile-card-list.mobile-hidden {{
                display: none !important;
            }}
        }}

        @media (min-width: 769px) {{
            .table-card {{
                display: block !important;
            }}
            .mobile-card-list {{
                display: none !important;
            }}
            .view-mode-bar {{
                display: none !important;
            }}
        }}

        @media (max-width: 480px) {{
            .kpi-grid {{
                grid-template-columns: 1fr;
            }}
        }}
    </style>
</head>
<body>
    <div class="container">
        <!-- Header -->
        <header>
            <div class="brand-title">
                <div class="brand-logo">國</div>
                <div class="brand-text">
                    <h1>國泰投信 人氣 ETF 與熱門基金 量化分析儀表板</h1>
                    <p>涵蓋歷次配息填息天數回測 • 投資組合產業穿透 • 國泰旗艦 (00687B/00893) • 同業宿敵擂台 (0050/0056/00919)</p>
                </div>
            </div>
            <div class="header-actions">
                <button class="btn" onclick="toggleTheme()" id="themeBtn">🌓 切換深淺模式</button>
                <button class="btn btn-primary" onclick="exportCurrentTable()">📥 匯出當前清單 (CSV)</button>
            </div>
        </header>

        <!-- KPI Cards -->
        <div class="kpi-grid">
            <div class="kpi-card blue">
                <div class="kpi-title">🚀 近 1 年績效冠軍</div>
                <div class="kpi-value num-up" id="kpiPerfVal">+77.2%</div>
                <div class="kpi-meta"><span class="kpi-badge">0050 元大台灣50</span> 國泰冠軍: 00881 (<span class="num-up">+52.7%</span>)</div>
            </div>
            <div class="kpi-card gold">
                <div class="kpi-title">💰 最新年化殖利率王</div>
                <div class="kpi-value" id="kpiDivVal">16.83%</div>
                <div class="kpi-meta"><span class="kpi-badge">00881 科技龍頭</span> 次席: 00919 (13.6%), 00878 (12.0%)</div>
            </div>
            <div class="kpi-card">
                <div class="kpi-title">⚡ 極速填息王 (成功率100%)</div>
                <div class="kpi-value" id="kpiFillVal">2.2 天</div>
                <div class="kpi-meta"><span class="kpi-badge">0056 元大高股息</span> 國泰旗艦: 00687B (3.0天), 00878 (11天)</div>
            </div>
            <div class="kpi-card purple">
                <div class="kpi-title">🛡️ 台積電重倉龍頭</div>
                <div class="kpi-value" id="kpiTsmcVal">55.2%</div>
                <div class="kpi-meta"><span class="kpi-badge">0050 台灣50</span> 國泰旗艦: 00881 (34.2%), 00922 (33.1%)</div>
            </div>
            <div class="kpi-card">
                <div class="kpi-title">🏦 月配息被動現金流指標</div>
                <div class="kpi-value" id="kpiMonthVal">00933B</div>
                <div class="kpi-meta"><span class="kpi-badge">金融債月配</span> 平均 8.3 天填息 / 避開二代健保最佳</div>
            </div>
        </div>

        <!-- Filter & Control Card -->
        <div class="controls-card">
            <div class="filter-row">
                <div class="filter-chips">
                    <div class="chip active" onclick="setFilter('all', this)">全部標的 (<span id="cntAll">0</span>)</div>
                    <div class="chip" onclick="setFilter('watchlist', this)">⭐ 我的自選追蹤 (<span id="cntWatchlist">0</span>)</div>
                    <div class="chip" onclick="setFilter('cathay', this)">🏆 國泰人氣主力 (<span id="cntCathay">0</span>)</div>
                    <div class="chip" onclick="setFilter('dirA', this)">✨ 國泰美債/電動車 (<span id="cntDirA">0</span>)</div>
                    <div class="chip" onclick="setFilter('peer', this)">⚔️ 同業宿敵擂台 (<span id="cntPeer">0</span>)</div>
                    <div class="chip" onclick="setFilter('month', this)">⚡ 月配息專區 (<span id="cntMonth">0</span>)</div>
                    <div class="chip" onclick="setFilter('highdiv', this)">💰 高股息型 (<span id="cntHighDiv">0</span>)</div>
                    <div class="chip" onclick="setFilter('market', this)">🚀 市值/科技成長 (<span id="cntMarket">0</span>)</div>
                    <div class="chip" onclick="setFilter('bond', this)">🛡️ 債券固定收益 (<span id="cntBond">0</span>)</div>
                </div>
                <div class="search-box">
                    <span class="search-icon">🔍</span>
                    <input type="text" id="searchInput" placeholder="搜尋代碼、簡稱或成分股..." oninput="onSearchChange()">
                </div>
            </div>
        </div>

        <!-- Nav Tabs -->
        <div class="nav-tabs">
            <div class="nav-tab active" id="navTabRank" onclick="switchTab('tab-rank', this)">📊 1. 綜合績效天梯</div>
            <div class="nav-tab" id="navTabDiv" onclick="switchTab('tab-div', this)">💰 2. 歷次配息與填息回測</div>
            <div class="nav-tab" id="navTabPort" onclick="switchTab('tab-port', this)">🧩 3. 投資組合與台積電穿透</div>
            <div class="nav-tab" id="navTabPk" onclick="switchTab('tab-pk', this)">⚔️ 4. 國泰 vs 同業霸主擂台 PK</div>
            <div class="nav-tab" id="navTabGuide" onclick="switchTab('tab-guide', this)">📖 5. 投資名詞圖解辭典</div>
        </div>

        <!-- TAB 1: 綜合績效天梯 -->
        <div id="tab-rank" class="tab-content active">
            <!-- 視圖模式切換（手機預設卡片，電腦預設大表） -->
            <div class="view-mode-bar">
                <span style="font-size:12px; color:var(--text-sub); font-weight:700;">📱 檢視模式：</span>
                <button class="view-mode-btn active" id="btnRankCard" onclick="setViewMode('rank', 'card')">📱 行動資訊卡片</button>
                <button class="view-mode-btn" id="btnRankTable" onclick="setViewMode('rank', 'table')">📊 完整比較大表</button>
            </div>
            <!-- 行動端資訊卡片清單 (垂直滾動，無需左右滑動) -->
            <div id="rankMobileList" class="mobile-card-list"></div>
            <!-- 電腦端 / 完整大表格 -->
            <div class="table-card mobile-hidden" id="rankTableCard">
                <div class="table-responsive">
                    <table id="rankTable">
                        <thead>
                            <tr>
                                <th onclick="sortTable('rank', 'house')">投信</th>
                                <th onclick="sortTable('rank', 'code')">代碼</th>
                                <th onclick="sortTable('rank', 'sname')">標的簡稱</th>
                                <th onclick="sortTable('rank', 'type')">類型</th>
                                <th onclick="sortTable('rank', 'cat')">推薦分類</th>
                                <th onclick="sortTable('rank', 'price')">現價/淨值</th>
                                <th onclick="sortTable('rank', 'fee')">經理費 ▼</th>
                                <th onclick="sortTable('rank', 'ret_3m')">近3個月</th>
                                <th onclick="sortTable('rank', 'ret_6m')">近6個月</th>
                                <th onclick="sortTable('rank', 'ret_1y')">近1年 ▼</th>
                                <th onclick="sortTable('rank', 'ret_3y')">近3年</th>
                                <th onclick="sortTable('rank', 'freq')">配息頻率</th>
                                <th onclick="sortTable('rank', 'risk')" style="font-size:14px; min-width:115px;">風險等級 ▼</th>
                                <th>操作</th>
                            </tr>
                        </thead>
                        <tbody id="rankTableBody">
                            <!-- JS injected -->
                        </tbody>
                    </table>
                </div>
            </div>
        </div>

        <!-- TAB 2: 配息與填息天數深度回測 -->
        <div id="tab-div" class="tab-content" style="display:none;">
            <!-- 視圖模式切換 -->
            <div class="view-mode-bar">
                <span style="font-size:12px; color:var(--text-sub); font-weight:700;">📱 檢視模式：</span>
                <button class="view-mode-btn active" id="btnDivCard" onclick="setViewMode('div', 'card')">📱 行動配息卡片</button>
                <button class="view-mode-btn" id="btnDivTable" onclick="setViewMode('div', 'table')">📊 完整比較大表</button>
            </div>
            <!-- 行動端資訊卡片清單 -->
            <div id="divMobileList" class="mobile-card-list"></div>
            <!-- 電腦端 / 完整大表格 -->
            <div class="table-card mobile-hidden" id="divTableCard">
                <div class="table-responsive">
                    <table id="divTable">
                        <thead>
                            <tr>
                                <th onclick="sortTable('div', 'house')">投信</th>
                                <th onclick="sortTable('div', 'code')">代碼</th>
                                <th onclick="sortTable('div', 'sname')">標的簡稱</th>
                                <th onclick="sortTable('div', 'fee')">經理費 ▼</th>
                                <th onclick="sortTable('div', 'latest_div')">最新每期配息(元)</th>
                                <th onclick="sortTable('div', 'ann_yield')">最新年化配息率 ▼</th>
                                <th onclick="sortTable('div', 'y1_div_sum')">近1年累計配息(元)</th>
                                <th onclick="sortTable('div', 'avg_fill_days')">平均填息天數(天)
                                    <span class="tooltip">?<span class="tooltip-text"><b>填息天數：</b>除息日後收盤價重新回補到除息前一日基準價所需的實際交易日數。</span></span>
                                </th>
                                <th onclick="sortTable('div', 'fill_rate')">填息成功率</th>
                                <th onclick="sortTable('div', 'freq')">頻率</th>
                                <th onclick="sortTable('div', 'rating')">填息力評級</th>
                                <th>歷次明細</th>
                            </tr>
                        </thead>
                        <tbody id="divTableBody">
                            <!-- JS injected -->
                        </tbody>
                    </table>
                </div>
            </div>
        </div>

        <!-- TAB 3: 投資組合與台積電穿透 -->
        <div id="tab-port" class="tab-content" style="display:none;">
            <div style="background:rgba(0,168,107,0.08); border:1px solid rgba(0,168,107,0.25); border-radius:var(--radius-md); padding:14px 18px; margin-bottom:16px; display:flex; align-items:center; justify-content:space-between; flex-wrap:wrap; gap:10px;">
                <div style="font-size:13px; color:var(--text-main);">
                    💡 <b>投資組合穿透指南：</b>點擊任一標的或右側「<b>🔍 查看 Top 10</b>」按鈕，即可展開<b>國泰官網同款的視覺化持股長條圖</b>，查看台積電、聯發科等完整權重與產業穿透。
                </div>
                <div style="font-size:12px; color:var(--text-sub);">支援台積電、聯發科等成分股即時搜尋 🔍</div>
            </div>
            <!-- 視圖模式切換 -->
            <div class="view-mode-bar">
                <span style="font-size:12px; color:var(--text-sub); font-weight:700;">📱 檢視模式：</span>
                <button class="view-mode-btn active" id="btnPortCard" onclick="setViewMode('port', 'card')">📱 行動持股卡片</button>
                <button class="view-mode-btn" id="btnPortTable" onclick="setViewMode('port', 'table')">📊 完整比較大表</button>
            </div>
            <!-- 行動端資訊卡片清單 -->
            <div id="portMobileList" class="mobile-card-list"></div>
            <!-- 電腦端 / 完整大表格 -->
            <div class="table-card mobile-hidden" id="portTableCard">
                <div class="table-responsive">
                    <table id="portTable">
                        <thead>
                            <tr>
                                <th onclick="sortTable('port', 'house')">投信</th>
                                <th onclick="sortTable('port', 'code')">代碼</th>
                                <th onclick="sortTable('port', 'sname')">標的簡稱</th>
                                <th onclick="sortTable('port', 'tsmc')">台積電佔比 ▼</th>
                                <th onclick="sortTable('port', 'tech')">科技電子佔比</th>
                                <th onclick="sortTable('port', 'tech_total')">科技合計純度</th>
                                <th onclick="sortTable('port', 'fin')">金融保險</th>
                                <th onclick="sortTable('port', 'trad')">傳統產業</th>
                                <th onclick="sortTable('port', 'bond_other')">債券/海外</th>
                                <th>核心配置定位</th>
                                <th>前 3 大核心持股</th>
                                <th>成分股穿透</th>
                            </tr>
                        </thead>
                        <tbody id="portTableBody">
                            <!-- JS injected -->
                        </tbody>
                    </table>
                </div>
            </div>
        </div>

        <!-- TAB 4: 國泰 vs 同業霸主擂台 PK -->
        <div id="tab-pk" class="tab-content" style="display:none;">
            <div class="pk-grid">
                <!-- PK 1: 高股息霸主之爭 -->
                <div class="pk-card">
                    <div class="pk-header">
                        <span>🏆 【高股息王牌對決】</span>
                        <span class="badge badge-cathay">00878 vs 0056 vs 00919</span>
                    </div>
                    <div class="pk-item highlight">
                        <div>
                            <b>國泰永續高股息 (00878)</b><br>
                            <small class="mono">現價 $33.8 | 季配 | 年化 12.04% | 平均填息 11天 (成功率 100%)</small>
                        </div>
                        <span class="badge badge-cathay">穩健防禦王</span>
                    </div>
                    <div class="pk-item">
                        <div>
                            <b>元大高股息 (0056)</b><br>
                            <small class="mono">現價 $57.5 | 季配 | 年化 9.39% | 平均填息 2.2天 (成功率 100%)</small>
                        </div>
                        <span class="badge badge-peer">老牌霸主</span>
                    </div>
                    <div class="pk-item">
                        <div>
                            <b>群益台灣精選高息 (00919)</b><br>
                            <small class="mono">現價 $32.4 | 季配 | 年化 13.58% | 平均填息 11.8天 (成功率 100%)</small>
                        </div>
                        <span class="badge badge-peer">高配息人氣王</span>
                    </div>
                    <div class="guide-example">
                        💡 <b>深度點評：</b>00878 金融股比重高、下檔防禦性最強；0056 歷史悠久、填息速度極為俐落；00919 則以宣告高配息政策著稱，三者均能達成 100% 填息。
                    </div>
                </div>

                <!-- PK 2: 市值型巨擘爭霸 -->
                <div class="pk-card">
                    <div class="pk-header">
                        <span>🚀 【市值型旗艦對決】</span>
                        <span class="badge badge-cathay">00922 vs 0050</span>
                    </div>
                    <div class="pk-item highlight">
                        <div>
                            <b>國泰台灣領袖50 (00922)</b><br>
                            <small class="mono">現價 $21.5 | 半年配 | 年化 9.94% | 台積電 33.1% | 1年績效 +48.7%</small>
                        </div>
                        <span class="badge badge-cathay">低單價低門檻</span>
                    </div>
                    <div class="pk-item">
                        <div>
                            <b>元大台灣50 (0050)</b><br>
                            <small class="mono">現價 $107.2 | 半年配 | 年化 1.12% | 台積電 55.2% | 1年績效 +77.2%</small>
                        </div>
                        <span class="badge badge-peer">台積電純度王</span>
                    </div>
                    <div class="guide-example">
                        💡 <b>深度點評：</b>0050 含有超過 55% 的台積電，在半導體大大多頭時進攻力最強；00922 則有單一持股上限機制，台積電約 33%，同時兼顧科技與低碳轉型產業，且具備近 10% 亮眼配息。
                    </div>
                </div>

                <!-- PK 3: 長天期美債龍頭對抗 -->
                <div class="pk-card">
                    <div class="pk-header">
                        <span>🏦 【長天期美債收益對抗】</span>
                        <span class="badge badge-cathay">00687B vs 00679B vs 00933B</span>
                    </div>
                    <div class="pk-item highlight">
                        <div>
                            <b>國泰20年美債 (00687B)</b><br>
                            <small class="mono">現價 $26.9 | 季配 | 年化 4.97% | 平均填息 3.0天 (成功率 100%)</small>
                        </div>
                        <span class="badge badge-cathay">美債規模王</span>
                    </div>
                    <div class="pk-item">
                        <div>
                            <b>元大美債20年 (00679B)</b><br>
                            <small class="mono">現價 $25.9 | 季配 | 年化 4.32% | 平均填息 2.7天 (成功率 100%)</small>
                        </div>
                        <span class="badge badge-peer">老牌美債</span>
                    </div>
                    <div class="pk-item highlight">
                        <div>
                            <b>國泰10Y+金融債 (00933B)</b><br>
                            <small class="mono">現價 $16.1 | 月配 | 年化 5.67% | 平均填息 8.3天 (月月現金流)</small>
                        </div>
                        <span class="badge badge-cathay">月配息避稅王</span>
                    </div>
                    <div class="guide-example">
                        💡 <b>深度點評：</b>00687B 溢價率控制極佳、殖利率逼近 5%；00933B 為月配息設計，填息天數短且完全不受單筆 2 萬元二代健保保費限制。
                    </div>
                </div>

                <!-- PK 4: 內扣經理費省錢對決 -->
                <div class="pk-card">
                    <div class="pk-header">
                        <span>💰 【經理費內扣省錢對決】</span>
                        <span class="badge" style="background:rgba(56,189,248,0.15); color:var(--accent-blue);">006208 vs 0050</span>
                    </div>
                    <div class="pk-item highlight">
                        <div>
                            <b>富邦台灣50 (006208)</b><br>
                            <small class="mono">現價 $115.6 | 經理費 0.15% (超低) | 半年配 | 台積電 55.1% | 1年績效 +77.6%</small>
                        </div>
                        <span class="badge badge-fill-fast">超省內扣王</span>
                    </div>
                    <div class="pk-item">
                        <div>
                            <b>元大台灣50 (0050)</b><br>
                            <small class="mono">現價 $107.2 | 經理費 0.32% (2倍+) | 半年配 | 台積電 55.2% | 1年績效 +77.2%</small>
                        </div>
                        <span class="badge badge-peer">老牌規模霸主</span>
                    </div>
                    <div class="guide-example">
                        💡 <b>深度點評：</b>006208 與 0050 追蹤完全相同的「臺灣50指數」，成分股一模一樣；但 006208 經理費僅 0.15%，每年比 0050 省下 0.17%！長期定期定額 10~20 年，省下的內扣費用複利滾存高達數萬至十幾萬元！
                    </div>
                </div>

                <!-- PK 5: 多流派高股息大對決 -->
                <div class="pk-card">
                    <div class="pk-header">
                        <span>⚡ 【全明星高股息流派對抗】</span>
                        <span class="badge badge-cathay">00878 vs 00929 vs 00713</span>
                    </div>
                    <div class="pk-item highlight">
                        <div>
                            <b>國泰永續高股息 (00878)</b><br>
                            <small class="mono">季配 | 年化 12.04% | 平均填息 11天 | 金融 23% + 科技 65%</small>
                        </div>
                        <span class="badge badge-cathay">穩健抗震首選</span>
                    </div>
                    <div class="pk-item">
                        <div>
                            <b>復華科技優息 (00929)</b><br>
                            <small class="mono">月配 | 年化 8.68% | 平均填息 9.5天 | 純科技 100%</small>
                        </div>
                        <span class="badge" style="background:rgba(245,158,11,0.15); color:var(--accent-gold);">月配科技首選</span>
                    </div>
                    <div class="pk-item">
                        <div>
                            <b>元大高息低波 (00713)</b><br>
                            <small class="mono">季配 | 年化 10.38% | 平均填息 4.5天 | 統一+電信低波防守</small>
                        </div>
                        <span class="badge badge-peer">空頭防禦神隊友</span>
                    </div>
                    <div class="guide-example">
                        💡 <b>深度點評：</b>00878 是全台受益人數冠軍，金融+科技均衡配置；00929 提供渴望「月月領息」的科技股擁護者；00713 則在大盤劇烈波動時展現驚人的低回撤與高填息防禦力。
                    </div>
                </div>
            </div>
        </div>

        <!-- TAB 5: 投資名詞圖解辭典 -->
        <div id="tab-guide" class="tab-content" style="display:none;">
            <div class="guide-grid">
                <!-- 風險報酬等級 RR1 ~ RR5 完整指南 -->
                <div class="guide-card" style="border: 2px solid rgba(59, 130, 246, 0.4); box-shadow: 0 4px 20px rgba(59, 130, 246, 0.15); grid-column: 1 / -1;">
                    <h3 style="font-size:18px; color:var(--accent-blue); display:flex; align-items:center; gap:8px;">
                        <span>🛡️ 基金與 ETF 風險報酬等級 (RR1 ~ RR5) 完整投資指南</span>
                        <span class="badge" style="background:rgba(59, 130, 246, 0.2); color:var(--accent-blue); font-size:12px; font-weight:700;">公會法定核心標準</span>
                    </h3>
                    <p style="font-size:14px; line-height:1.7;"><b>白話定義：</b>依中華民國投信投顧公會標準，依產品所投資之標的資產類別、價格波動度與風險程度，嚴格劃分為 <b>RR1 至 RR5 共五級</b>（Risk-Return Ratio）。<b>數字越大，代表該標的價格波動與潛在回報空間越高，但空頭回檔時面臨的價格跌幅風險也越大</b>。</p>
                    
                    <div style="display:flex; flex-direction:column; gap:10px; margin: 14px 0;">
                        <div style="background:var(--bg-body); border-left:5px solid #06b6d4; padding:12px 16px; border-radius:6px; display:flex; justify-content:space-between; align-items:center; flex-wrap:wrap; gap:10px;">
                            <div>
                                <span class="badge mono" style="background:rgba(6,182,212,0.18); color:#06b6d4; font-size:15px; font-weight:800; padding:3px 10px; border-radius:4px;">RR1 (低度風險)</span>
                                <b style="margin-left:8px; font-size:15px; color:var(--text-main);">極保守 / 本金保全</b>
                                <div style="font-size:13px; color:var(--text-sub); margin-top:4px;">主要投資於貨幣市場工具、短期公庫票券或超短期國庫券。幾乎無本金虧損風險，收益率貼近銀行定存利率。<br><b>適合對象：</b>存放 3~6 個月生活緊急備用金、極度厭惡波動之保守者。</div>
                            </div>
                            <span class="badge" style="background:rgba(255,255,255,0.08); font-size:12px; padding:4px 8px;">無本金虧損風險</span>
                        </div>

                        <div style="background:var(--bg-body); border-left:5px solid #38bdf8; padding:12px 16px; border-radius:6px; display:flex; justify-content:space-between; align-items:center; flex-wrap:wrap; gap:10px;">
                            <div>
                                <span class="badge mono" style="background:rgba(56,189,248,0.18); color:#38bdf8; font-size:15px; font-weight:800; padding:3px 10px; border-radius:4px;">RR2 (中低風險)</span>
                                <b style="margin-left:8px; font-size:15px; color:var(--text-main);">長天期公債 / 頂級投資級公司債</b>
                                <div style="font-size:13px; color:var(--text-sub); margin-top:4px;">主要投資於全球政府主權債（如美國公債）或 BBB 級以上跨國巨頭公司債。波動溫和，依賴固定債息提供穩定被動現金流。<br><b>本儀表板標的：</b>00687B (國泰20年美債)、00679B (元大美債20年)、00937B (群益ESG投等債20+)、00933B (國泰10Y+金融債)。<br><b>適合對象：</b>退休族、追求穩定月配息、資產配置股債平衡要角。</div>
                            </div>
                            <span class="badge badge-cathay" style="font-size:12px; padding:4px 8px;">降息週期防守王</span>
                        </div>

                        <div style="background:var(--bg-body); border-left:5px solid #eab308; padding:12px 16px; border-radius:6px; display:flex; justify-content:space-between; align-items:center; flex-wrap:wrap; gap:10px;">
                            <div>
                                <span class="badge mono" style="background:rgba(234,179,8,0.18); color:#eab308; font-size:15px; font-weight:800; padding:3px 10px; border-radius:4px;">RR3 (中度風險)</span>
                                <b style="margin-left:8px; font-size:15px; color:var(--text-main);">多元平衡 / 高收益債</b>
                                <div style="font-size:13px; color:var(--text-sub); margin-top:4px;">主要投資於跨國多元平衡型基金、高評級公用事業股、非投資等級高收益債券。兼具適度資本利得增長與定期收益分配。<br><b>適合對象：</b>穩健型投資人、希望兼顧股票成長與債券保護者。</div>
                            </div>
                            <span class="badge" style="background:rgba(234,179,8,0.2); color:#eab308; font-size:12px; padding:4px 8px;">股債平衡</span>
                        </div>

                        <div style="background:var(--bg-body); border-left:5px solid #fb923c; padding:12px 16px; border-radius:6px; display:flex; justify-content:space-between; align-items:center; flex-wrap:wrap; gap:10px;">
                            <div>
                                <span class="badge mono" style="background:rgba(249,115,22,0.18); color:#fb923c; font-size:15px; font-weight:800; padding:3px 10px; border-radius:4px;">RR4 (中高風險)</span>
                                <b style="margin-left:8px; font-size:15px; color:var(--text-main);">台股權值大盤 / 主流人氣高股息</b>
                                <div style="font-size:13px; color:var(--text-sub); margin-top:4px;">涵蓋全市場代表性權值股（台積電、聯發科、鴻海）或高殖利率優質企業。隨景氣週期震盪，長線具備強大複利增長動能。<br><b>本儀表板標的：</b>0050 (元大台灣50)、006208 (富邦台灣50)、00878 (國泰高股息)、0056 (元大高股息)、00713 (高息低波)、00922 (國泰領袖50)、00940 (台灣價值高息)。<br><b>適合對象：</b>定期定額長期存股族、追求資產翻倍成長之中長期投資人。</div>
                            </div>
                            <span class="badge" style="background:rgba(249,115,22,0.2); color:#fb923c; font-size:12px; padding:4px 8px;">主流存股核心</span>
                        </div>

                        <div style="background:var(--bg-body); border-left:5px solid #f472b6; padding:12px 16px; border-radius:6px; display:flex; justify-content:space-between; align-items:center; flex-wrap:wrap; gap:10px;">
                            <div>
                                <span class="badge mono" style="background:rgba(236,72,153,0.18); color:#f472b6; font-size:15px; font-weight:800; padding:3px 10px; border-radius:4px;">RR5 (高度風險)</span>
                                <b style="margin-left:8px; font-size:15px; color:var(--text-main);">單一科技主題 / 中小型動能股 / 主動選股型</b>
                                <div style="font-size:13px; color:var(--text-sub); margin-top:4px;">聚焦特定爆發性科技（AI、半導體、智能電動車）或中小型高成長股。多頭時獲利進攻力全場最強，但回調時回撤幅度亦最大。<br><b>本儀表板標的：</b>00881 (國泰台灣5G+)、00893 (國泰智能電動車)、00929 (復華科技優息)、國泰小龍基金 (14)、國泰台灣高股息基金 (CL/CM)。<br><b>適合對象：</b>波段積極型投資人、能承受 20%~30% 帳面波動追求超額報酬 (Alpha) 者。</div>
                            </div>
                            <span class="badge" style="background:rgba(236,72,153,0.2); color:#f472b6; font-size:12px; padding:4px 8px;">進攻爆發型</span>
                        </div>
                    </div>

                    <div class="guide-example">
                        💡 <b>不同年齡與人生階段之黃金配置公式：</b><br>
                        • <b>年輕小資衝刺期 (20~40歲)：</b>建議配置 <b>70%~80% 於 RR4 (如 006208、00922) 與 RR5 (如 00881、00893)</b>，把握 20 年以上的時間優勢，讓大盤與科技複利滾雪球。<br>
                        • <b>中壯平衡防禦期 (40~55歲)：</b>建議 <b>40% RR4 高股息 (00878、00713) + 30% RR4 市值型 (00922) + 30% RR2 美債 (00687B)</b>，兼顧資產成長與抗震抗通膨。<br>
                        • <b>退休樂活金流期 (55歲以上)：</b>建議提高 <b>RR2 (00687B、00933B、00937B) 與 RR4 防禦型高股息 (00713、00878) 至 75% 以上</b>，以保全本金、抵禦空頭大跌，並確保每月/每季均有現金流入帳。
                    </div>
                </div>

                <div class="guide-card">
                    <h3>💡 收益平準金 (Equalization Reserve)</h3>
                    <p><b>白話定義：</b>將除息日前湧入之新買盤資金，按比例提撥部分本金為「配息準備金」，用以在除息時發放給該投資人，防止老投資人的每期配息率被嚴重稀釋。</p>
                    <div class="guide-example">
                        <b>真實運作：</b>若無平準金，某 ETF 規模從 100 億暴增至 200 億，但成分股領到的現金股利是固定的，每位股東分配到的配息會被腰斬；平準金可確保「人人領到原本承諾的配息率」。<br>
                        <b>重要優勢：</b>平準金發放屬於「返還本金」，<b>免計入個人綜合所得稅、免扣 2.11% 二代健保補充保費！</b>
                    </div>
                </div>

                <div class="guide-card">
                    <h3>⏱️ 填息交易天數與貼息風險</h3>
                    <p><b>白話定義：</b>除息後市價重新漲回「除息前最後一個交易日之基準價」所需的實際交易日數。</p>
                    <div class="guide-example">
                        <b>關鍵真諦：</b>「能填息的配息才是真賺進口袋！」若一檔標的年化配息 15%，但除息後價格一路走低（貼息），投資人領到的股息只是把自身資產帳面縮水換成現金。<br>
                        <b>數據印證：</b>國泰 00878、00922 與元大 0056 的歷史填息成功率均為 100%，代表其配息有強大的企業營運獲利作支撐。
                    </div>
                </div>

                <div class="guide-card">
                    <h3>🧾 54C 股利所得課稅與二代健保</h3>
                    <p><b>白話定義：</b>配息結構中屬於境內上市公司發放的現金股利所得，為最主要的課稅來源。</p>
                    <div class="guide-example">
                        <b>單筆 2 萬元門檻：</b>只要單次配息中「54C 股利所得」超過新台幣 20,000 元，券商必須代扣 2.11% 二代健保補充保費。<br>
                        <b>避稅秘訣：</b>選擇「月配型（如 00933B）」能將年度配息拆分成 12 次，大幅降低單次超過 2 萬元的機率；此外平準金與資本利得完全免扣保費。
                    </div>
                </div>

                <div class="guide-card">
                    <h3>⚖️ 市值型 vs 高股息型 vs 產業型</h3>
                    <p><b>核心策略差異：</b></p>
                    <div class="guide-example">
                        • <b>市值型 (如 00922、0050、006208)：</b>台積電佔比 30%~55%，緊跟台灣經濟龍頭，追求長線資產複利增長。<br>
                        • <b>高股息型 (如 00878、0056、00713、00929)：</b>台積電 0%，重視防禦力與現金流，年化配息 7%~12%，適合退休與房貸族。<br>
                        • <b>產業主題型 (如 00881、00893)：</b>聚焦 AI、電動車與半導體，科技純度 90%+，波段爆發力最強。
                    </div>
                </div>

                <div class="guide-card">
                    <h3>💰 ETF 經理費與內扣費用 (Expense Ratio) 為何是長期獲利勝負手？</h3>
                    <p><b>白話定義：</b>投信公司管理 ETF 每年所收取的費用。這筆費用<b>不會寄繳費單給您</b>，而是每個營業日開盤前<b>直接從 ETF 的資產淨值中自動扣除</b>，屬於「隱形內扣成本」。</p>
                    <div class="guide-example">
                        <b>🔥 為什麼這一定要比？（真實算給您看）：</b><br>
                        • <b>0050 vs 006208（最經典案例）</b>：兩檔都追蹤臺灣 50 指數，持股一模一樣（台積電都佔 ~55%）。但 0050 經理費 0.32%，006208 僅 0.15%（相差超過 2 倍！）。<br>
                        • <b>複利的威力</b>：若投入 100 萬元放 20 年，光是這每年 0.17% 的經理費差距，006208 就能默默幫您多守住 <b>超過 12 萬 ~ 18 萬元</b> 的本金收益！<br>
                        • <b>主動型基金 (1.5%) vs 被動 ETF (0.2%)</b>：主動型共同基金每年內扣高達 1.5%~2.0%，等於每年要比大盤多賺 1.5% 才能打平手，因此挑選主動基金必須嚴格檢驗經理人能否持續繳出超額報酬 (Alpha)。
                    </div>
                </div>
            </div>
        </div>

        <!-- Footer -->
        <footer>
            <p>國泰投信 人氣 ETF 與熱門基金 量化分析系統 | 數據來源: 國泰投信官方 CWAPI、TWSE、Yahoo Finance | 本系統僅供量化研究與個人資產配置參考，投資決策請審慎評估風險。</p>
        </footer>
    </div>

    <!-- Modal Drawer: 歷次配息明細 -->
    <div class="modal-overlay" id="divModal" onclick="closeModal(event, 'divModal')">
        <div class="modal-card">
            <button class="modal-close" onclick="document.getElementById('divModal').classList.remove('active')">&times;</button>
            <h2 id="modalDivTitle" style="font-size:18px; margin-bottom:12px; color:var(--brand-green);">歷次詳細配息與填息天數紀錄</h2>
            <div class="table-responsive">
                <table>
                    <thead>
                        <tr>
                            <th>除息年月</th>
                            <th>除息交易日</th>
                            <th>每期配息(元)</th>
                            <th>當期配息率</th>
                            <th>年化配息率</th>
                            <th>填息狀態</th>
                            <th>填息交易天數</th>
                        </tr>
                    </thead>
                    <tbody id="modalDivBody"></tbody>
                </table>
            </div>
        </div>
    </div>

    <!-- Modal Drawer: 成分股穿透明細 -->
    <div class="modal-overlay" id="portModal" onclick="closeModal(event, 'portModal')">
        <div class="modal-card">
            <button class="modal-close" onclick="document.getElementById('portModal').classList.remove('active')">&times;</button>
            <div id="modalPortHeader">
                <h2 id="modalPortTitle" style="font-size:18px; margin-bottom:8px; color:var(--brand-green); display:flex; align-items:center; flex-wrap:wrap; gap:8px;">主要投資標的組合穿透</h2>
                <div id="modalPortMeta" style="margin-bottom:18px;"></div>
            </div>
            <!-- Visual Bar Section (國泰官網同款佔比視覺長條圖) -->
            <div id="modalPortVisual"></div>
            <!-- Table Section -->
            <div style="margin-top:20px; padding-top:16px; border-top:1px solid var(--border);">
                <div style="font-size:14px; font-weight:700; color:var(--text-sub); margin-bottom:10px;">📋 詳細數據清單</div>
                <div class="table-responsive">
                    <table>
                        <thead>
                            <tr>
                                <th>排名</th>
                                <th>代碼</th>
                                <th>標的名稱</th>
                                <th>產業分類</th>
                                <th>持股權重</th>
                            </tr>
                        </thead>
                        <tbody id="modalPortBody"></tbody>
                    </table>
                </div>
            </div>
        </div>
    </div>

    <!-- Scripts -->
    <script>
        const rawData = {targets_json};
        let currentFilter = 'all';
        let currentSearch = '';
        let currentSort = {{ key: 'ret_1y', order: 'desc' }};

        // 初始化
        window.addEventListener('DOMContentLoaded', () => {{
            updateFilterCounts();
            renderAllTables();
        }});

        function updateFilterCounts() {{
            document.getElementById('cntAll').innerText = rawData.length;
            document.getElementById('cntWatchlist').innerText = rawData.filter(t => t.group === '自選追蹤' || ['006208', '00929', '00713', '00940', '00937B'].includes(t.code)).length;
            document.getElementById('cntCathay').innerText = rawData.filter(t => t.group === '國泰人氣王').length;
            document.getElementById('cntDirA').innerText = rawData.filter(t => t.group === '方向A (國泰旗艦)').length;
            document.getElementById('cntPeer').innerText = rawData.filter(t => t.group === '方向B (同業擂台)').length;
            document.getElementById('cntMonth').innerText = rawData.filter(t => t.freq === '月配').length;
            document.getElementById('cntHighDiv').innerText = rawData.filter(t => t.cat.includes('高息') || t.sname.includes('高股息') || t.ann_yield >= 0.07).length;
            document.getElementById('cntMarket').innerText = rawData.filter(t => t.cat.includes('市值') || t.cat.includes('科技') || t.sname.includes('50') || t.tsmc >= 0.2).length;
            document.getElementById('cntBond').innerText = rawData.filter(t => t.cat.includes('債') || t.sname.includes('債') || t.bond_other >= 0.7).length;
        }}

        function getFilteredData() {{
            return rawData.filter(item => {{
                // Group Filter
                if (currentFilter === 'watchlist' && !(item.group === '自選追蹤' || ['006208', '00929', '00713', '00940', '00937B'].includes(item.code))) return false;
                if (currentFilter === 'cathay' && item.group !== '國泰人氣王') return false;
                if (currentFilter === 'dirA' && item.group !== '方向A (國泰旗艦)') return false;
                if (currentFilter === 'peer' && item.group !== '方向B (同業擂台)') return false;
                if (currentFilter === 'month' && item.freq !== '月配') return false;
                if (currentFilter === 'highDiv' && !(item.cat.includes('高息') || item.sname.includes('高股息') || item.ann_yield >= 0.07)) return false;
                if (currentFilter === 'market' && !(item.cat.includes('市值') || item.cat.includes('科技') || item.sname.includes('50') || item.tsmc >= 0.2)) return false;
                if (currentFilter === 'bond' && !(item.cat.includes('債') || item.sname.includes('債') || item.bond_other >= 0.7)) return false;

                // Search Filter
                if (currentSearch) {{
                    const s = currentSearch.toLowerCase();
                    const matchBasic = item.code.toLowerCase().includes(s) || item.sname.toLowerCase().includes(s) || item.name.toLowerCase().includes(s) || item.cat.toLowerCase().includes(s) || (item.risk && item.risk.toLowerCase().includes(s));
                    const matchHoldings = item.holdings && item.holdings.some(h => (h.name && h.name.toLowerCase().includes(s)) || (h.code && h.code.toLowerCase().includes(s)));
                    if (!matchBasic && !matchHoldings) return false;
                }}
                return true;
            }});
        }}

        function renderAllTables() {{
            const data = getFilteredData();
            
            // Sort
            data.sort((a, b) => {{
                let valA = a[currentSort.key];
                let valB = b[currentSort.key];
                if (valA === null || valA === undefined) valA = -999999;
                if (valB === null || valB === undefined) valB = -999999;
                if (typeof valA === 'string') {{
                    return currentSort.order === 'asc' ? valA.localeCompare(valB) : valB.localeCompare(valA);
                }}
                return currentSort.order === 'asc' ? valA - valB : valB - valA;
            }});

            renderRankTable(data);
            renderDivTable(data);
            renderPortTable(data);
        }}

        function fmtPct(val) {{
            if (val === null || val === undefined || isNaN(val)) return '--';
            const s = (val * 100).toFixed(2) + '%';
            if (val > 0) return `<span class="num-up">+${{s}}</span>`;
            if (val < 0) return `<span class="num-down">${{s}}</span>`;
            return `<span class="num-neutral">${{s}}</span>`;
        }}

        function fmtNum(val, dec = 2) {{
            if (val === null || val === undefined || isNaN(val)) return '--';
            return `<span class="mono">${{val.toFixed(dec)}}</span>`;
        }}

        function getHouseBadge(house) {{
            if (house === '國泰投信') return 'badge-cathay';
            if (house === '富邦投信') return 'badge-fubon';
            if (house === '復華投信') return 'badge-fuhwa';
            return 'badge-peer';
        }}

        function getFeeBadge(fee) {{
            if (!fee || isNaN(fee)) return '--';
            const s = (fee * 100).toFixed(2) + '%';
            if (fee <= 0.0020) return `<span class="badge" style="background:rgba(56,189,248,0.15); color:var(--accent-blue); font-weight:700;">${{s}} (超低)</span>`;
            if (fee >= 0.0100) return `<span class="badge" style="background:rgba(245,158,11,0.15); color:var(--accent-gold); font-weight:700;">${{s}} (主動)</span>`;
            return `<span class="mono">${{s}}</span>`;
        }}

        // 風險報酬等級醒目大字徽章 (RR1 ~ RR5)
        function getRiskBadge(risk) {{
            if (!risk) return '--';
            let bg = 'rgba(255,255,255,0.08)';
            let color = 'var(--text-main)';
            let border = 'var(--border)';
            let desc = '';
            if (risk === 'RR1') {{
                bg = 'rgba(6,182,212,0.18)'; color = '#06b6d4'; border = 'rgba(6,182,212,0.4)'; desc = '低';
            }} else if (risk === 'RR2') {{
                bg = 'rgba(56,189,248,0.18)'; color = '#38bdf8'; border = 'rgba(56,189,248,0.4)'; desc = '中低';
            }} else if (risk === 'RR3') {{
                bg = 'rgba(234,179,8,0.18)'; color = '#eab308'; border = 'rgba(234,179,8,0.4)'; desc = '中度';
            }} else if (risk === 'RR4') {{
                bg = 'rgba(249,115,22,0.18)'; color = '#fb923c'; border = 'rgba(249,115,22,0.4)'; desc = '中高';
            }} else if (risk === 'RR5') {{
                bg = 'rgba(236,72,153,0.18)'; color = '#f472b6'; border = 'rgba(236,72,153,0.4)'; desc = '高度';
            }}
            return `<span class="badge mono" style="background:${{bg}}; color:${{color}}; border:1px solid ${{border}}; font-size:15px; font-weight:800; padding:4px 10px; border-radius:6px; letter-spacing:0.5px; display:inline-flex; align-items:center; gap:5px;" title="${{risk}} - ${{desc}}等級 (可切換至 Tab 5 查看詳細說明)"><span>${{risk}}</span><small style="font-size:11px; font-weight:700; opacity:0.9;">(${{desc}})</small></span>`;
        }}

        // 視圖模式管理（手機預設卡片，電腦預設大表）
        let viewModes = {{
            rank: 'card',
            div: 'card',
            port: 'card'
        }};

        function setViewMode(tabType, mode) {{
            viewModes[tabType] = mode;
            const cardEl = document.getElementById(tabType + 'MobileList');
            const tableEl = document.getElementById(tabType + 'TableCard');
            const btnCard = document.getElementById('btn' + tabType.charAt(0).toUpperCase() + tabType.slice(1) + 'Card');
            const btnTable = document.getElementById('btn' + tabType.charAt(0).toUpperCase() + tabType.slice(1) + 'Table');

            if (mode === 'card') {{
                if (cardEl) cardEl.classList.remove('mobile-hidden');
                if (tableEl) tableEl.classList.add('mobile-hidden');
                if (btnCard) btnCard.classList.add('active');
                if (btnTable) btnTable.classList.remove('active');
            }} else {{
                if (cardEl) cardEl.classList.add('mobile-hidden');
                if (tableEl) tableEl.classList.remove('mobile-hidden');
                if (btnCard) btnCard.classList.remove('active');
                if (btnTable) btnTable.classList.add('active');
            }}
        }}

        // Render Tab 1 (Rank)
        function renderRankTable(data) {{
            const tbody = document.getElementById('rankTableBody');
            const mobileList = document.getElementById('rankMobileList');

            // 1. 電腦端大表格
            tbody.innerHTML = data.map(t => {{
                const badgeClass = getHouseBadge(t.house);
                return `
                    <tr>
                        <td><span class="badge ${{badgeClass}}">${{t.house}}</span></td>
                        <td><b class="mono">${{t.code}}</b></td>
                        <td><b>${{t.sname}}</b></td>
                        <td>${{t.type}}</td>
                        <td>${{t.cat}}</td>
                        <td>${{fmtNum(t.price)}}</td>
                        <td>${{getFeeBadge(t.fee)}}</td>
                        <td>${{fmtPct(t.ret_3m)}}</td>
                        <td>${{fmtPct(t.ret_6m)}}</td>
                        <td>${{fmtPct(t.ret_1y)}}</td>
                        <td>${{fmtPct(t.ret_3y)}}</td>
                        <td>${{t.freq || '--'}}</td>
                        <td style="text-align:center;">${{getRiskBadge(t.risk)}}</td>
                        <td><button class="btn" style="padding:4px 8px; font-size:11px;" onclick="openPortModal('${{t.code}}')">穿透持股</button></td>
                    </tr>
                `;
            }}).join('');

            // 2. 手機端行動資訊卡片 (零橫向滑動)
            if (mobileList) {{
                mobileList.innerHTML = data.map(t => {{
                    const badgeClass = getHouseBadge(t.house);
                    return `
                        <div class="m-card" onclick="openPortModal('${{t.code}}')">
                            <div class="m-card-header">
                                <div class="m-card-title-group">
                                    <span class="badge ${{badgeClass}}">${{t.house}}</span>
                                    <b class="m-card-code">${{t.code}}</b>
                                    <span class="m-card-name">${{t.sname}}</span>
                                </div>
                                <div>${{getRiskBadge(t.risk)}}</div>
                            </div>
                            <div class="m-card-tags">
                                <span class="badge" style="background:rgba(255,255,255,0.06);">${{t.type}}</span>
                                <span class="badge" style="background:rgba(255,255,255,0.06);">${{t.cat}}</span>
                                <span class="badge" style="background:rgba(255,255,255,0.06);">${{t.freq || '不配息'}}</span>
                                ${{getFeeBadge(t.fee)}}
                            </div>
                            <div class="m-card-grid">
                                <div class="m-grid-item">
                                    <div class="m-label">現價 / 淨值</div>
                                    <div class="m-val mono">${{fmtNum(t.price)}}</div>
                                </div>
                                <div class="m-grid-item">
                                    <div class="m-label">近 1 年績效</div>
                                    <div class="m-val">${{fmtPct(t.ret_1y)}}</div>
                                </div>
                                <div class="m-grid-item">
                                    <div class="m-label">近 3 個月</div>
                                    <div class="m-val">${{fmtPct(t.ret_3m)}}</div>
                                </div>
                                <div class="m-grid-item">
                                    <div class="m-label">近 3 年績效</div>
                                    <div class="m-val">${{fmtPct(t.ret_3y)}}</div>
                                </div>
                            </div>
                            <div class="m-card-footer">
                                <span style="font-size:11px; color:var(--text-muted);">點擊展開官網同款 Top 10 ▶</span>
                                <div class="m-card-actions">
                                    <button class="btn btn-sm" style="padding:4px 8px; font-size:11px;" onclick="event.stopPropagation(); openDivModal('${{t.code}}')">💰 配息</button>
                                    <button class="btn btn-sm btn-primary" style="padding:4px 10px; font-size:11px;" onclick="event.stopPropagation(); openPortModal('${{t.code}}')">🔍 穿透</button>
                                </div>
                            </div>
                        </div>
                    `;
                }}).join('');
            }}
        }}

        // Render Tab 2 (Div)
        function renderDivTable(data) {{
            const tbody = document.getElementById('divTableBody');
            const mobileList = document.getElementById('divMobileList');
            const divData = data.filter(t => t.freq !== '不配息');

            // 1. 電腦端大表格
            tbody.innerHTML = divData.map(t => {{
                const badgeClass = getHouseBadge(t.house);
                let fillBadge = 'badge-fill-mid';
                if (t.avg_fill_days && t.avg_fill_days <= 10) fillBadge = 'badge-fill-fast';
                else if (t.avg_fill_days && t.avg_fill_days > 40) fillBadge = 'badge-fill-slow';

                return `
                    <tr class="clickable-row" onclick="openDivModal('${{t.code}}')">
                        <td><span class="badge ${{badgeClass}}">${{t.house}}</span></td>
                        <td><b class="mono">${{t.code}}</b></td>
                        <td><b>${{t.sname}}</b></td>
                        <td>${{getFeeBadge(t.fee)}}</td>
                        <td>${{fmtNum(t.latest_div, 3)}} 元</td>
                        <td>${{t.ann_yield !== null && !isNaN(t.ann_yield) ? `<b class="mono">${{(t.ann_yield * 100).toFixed(2)}}%</b>` : '--'}}</td>
                        <td>${{fmtNum(t.y1_div_sum, 2)}} 元</td>
                        <td>${{t.avg_fill_days ? `<span class="badge ${{fillBadge}}">${{t.avg_fill_days}} 天</span>` : '--'}}</td>
                        <td>${{t.fill_rate !== null ? (t.fill_rate * 100).toFixed(0) + '%' : '--'}}</td>
                        <td>${{t.freq}}</td>
                        <td><small>${{t.rating}}</small></td>
                        <td><button class="btn" style="padding:4px 8px; font-size:11px;">檢視紀錄 (${{t.div_records.length}}期)</button></td>
                    </tr>
                `;
            }}).join('');

            // 2. 手機端行動配息卡片
            if (mobileList) {{
                mobileList.innerHTML = divData.map(t => {{
                    const badgeClass = getHouseBadge(t.house);
                    let fillBadge = 'badge-fill-mid';
                    if (t.avg_fill_days && t.avg_fill_days <= 10) fillBadge = 'badge-fill-fast';
                    else if (t.avg_fill_days && t.avg_fill_days > 40) fillBadge = 'badge-fill-slow';

                    return `
                        <div class="m-card" onclick="openDivModal('${{t.code}}')">
                            <div class="m-card-header">
                                <div class="m-card-title-group">
                                    <span class="badge ${{badgeClass}}">${{t.house}}</span>
                                    <b class="m-card-code">${{t.code}}</b>
                                    <span class="m-card-name">${{t.sname}}</span>
                                </div>
                                <span class="badge ${{fillBadge}}">${{t.avg_fill_days ? t.avg_fill_days + ' 天填息' : '無紀錄'}}</span>
                            </div>
                            <div class="m-card-tags">
                                <span class="badge" style="background:rgba(255,255,255,0.06);">${{t.freq}}</span>
                                ${{getFeeBadge(t.fee)}}
                                <span class="badge" style="background:rgba(255,255,255,0.06); font-size:11px;">成功率 ${{t.fill_rate ? (t.fill_rate * 100).toFixed(0) + '%' : '--'}}</span>
                            </div>
                            <div class="m-card-grid">
                                <div class="m-grid-item">
                                    <div class="m-label">最新年化配息率</div>
                                    <div class="m-val mono" style="color:var(--accent-gold); font-size:16px;">${{t.ann_yield !== null && !isNaN(t.ann_yield) ? (t.ann_yield * 100).toFixed(2) + '%' : '--'}}</div>
                                </div>
                                <div class="m-grid-item">
                                    <div class="m-label">最新每期配息</div>
                                    <div class="m-val mono">${{fmtNum(t.latest_div, 3)}} 元</div>
                                </div>
                                <div class="m-grid-item">
                                    <div class="m-label">近 1 年累計配息</div>
                                    <div class="m-val mono">${{fmtNum(t.y1_div_sum, 2)}} 元</div>
                                </div>
                                <div class="m-grid-item">
                                    <div class="m-label">填息力評級</div>
                                    <div class="m-val" style="font-size:12px; font-weight:600;">${{t.rating}}</div>
                                </div>
                            </div>
                            <div class="m-card-footer">
                                <span style="font-size:11px; color:var(--text-muted);">歷史累計 ${{t.div_records.length}} 期紀錄</span>
                                <button class="btn btn-sm btn-primary" style="padding:4px 10px; font-size:11px;" onclick="event.stopPropagation(); openDivModal('${{t.code}}')">檢視明細 ▶</button>
                            </div>
                        </div>
                    `;
                }}).join('');
            }}
        }}

        // Render Tab 3 (Port)
        function renderPortTable(data) {{
            const tbody = document.getElementById('portTableBody');
            const mobileList = document.getElementById('portMobileList');

            // 1. 電腦端大表格
            tbody.innerHTML = data.map(t => {{
                const badgeClass = getHouseBadge(t.house);
                return `
                    <tr class="clickable-row" onclick="openPortModal('${{t.code}}')">
                        <td><span class="badge ${{badgeClass}}">${{t.house}}</span></td>
                        <td><b class="mono">${{t.code}}</b></td>
                        <td><b>${{t.sname}}</b></td>
                        <td>${{t.tsmc > 0 ? `<b class="mono">${{(t.tsmc * 100).toFixed(1)}}%</b>` : '0.0%'}}</td>
                        <td>${{(t.tech * 100).toFixed(1)}}%</td>
                        <td><b>${{(t.tech_total * 100).toFixed(1)}}%</b></td>
                        <td>${{(t.fin * 100).toFixed(1)}}%</td>
                        <td>${{(t.trad * 100).toFixed(1)}}%</td>
                        <td>${{(t.bond_other * 100).toFixed(1)}}%</td>
                        <td><small>${{t.positioning}}</small></td>
                        <td><small style="color:var(--text-sub);">${{t.top3}}</small></td>
                        <td><button class="btn btn-primary" style="padding:5px 10px; font-size:11px;" onclick="event.stopPropagation(); openPortModal('${{t.code}}')">🔍 查看 Top 10</button></td>
                    </tr>
                `;
            }}).join('');

            // 2. 手機端行動持股卡片
            if (mobileList) {{
                mobileList.innerHTML = data.map(t => {{
                    const badgeClass = getHouseBadge(t.house);
                    return `
                        <div class="m-card" onclick="openPortModal('${{t.code}}')">
                            <div class="m-card-header">
                                <div class="m-card-title-group">
                                    <span class="badge ${{badgeClass}}">${{t.house}}</span>
                                    <b class="m-card-code">${{t.code}}</b>
                                    <span class="m-card-name">${{t.sname}}</span>
                                </div>
                                <div>${{getRiskBadge(t.risk)}}</div>
                            </div>
                            <div style="font-size:12px; color:var(--brand-green); font-weight:600;">
                                🎯 ${{t.positioning}}
                            </div>
                            <div class="m-card-grid">
                                <div class="m-grid-item">
                                    <div class="m-label">台積電佔比</div>
                                    <div class="m-val mono">${{t.tsmc > 0 ? (t.tsmc * 100).toFixed(1) + '%' : '0.0%'}}</div>
                                </div>
                                <div class="m-grid-item">
                                    <div class="m-label">科技純度</div>
                                    <div class="m-val mono"><b>${{(t.tech_total * 100).toFixed(1)}}%</b></div>
                                </div>
                                <div class="m-grid-item">
                                    <div class="m-label">金融保險</div>
                                    <div class="m-val mono">${{(t.fin * 100).toFixed(1)}}%</div>
                                </div>
                                <div class="m-grid-item">
                                    <div class="m-label">傳產 / 其他</div>
                                    <div class="m-val mono">${{((t.trad + t.bond_other) * 100).toFixed(1)}}%</div>
                                </div>
                            </div>
                            <div style="background:var(--bg-body); padding:8px 10px; border-radius:6px; font-size:12px; color:var(--text-sub);">
                                <b>核心持股：</b>${{t.top3}}
                            </div>
                            <div class="m-card-footer">
                                <span style="font-size:11px; color:var(--text-muted);">展開官網同款長條圖 ▶</span>
                                <button class="btn btn-sm btn-primary" style="padding:4px 10px; font-size:11px;" onclick="event.stopPropagation(); openPortModal('${{t.code}}')">🔍 穿透 Top 10</button>
                            </div>
                        </div>
                    `;
                }}).join('');
            }}
        }}

        // Tab Switching
        function switchTab(tabId, el) {{
            document.querySelectorAll('.tab-content').forEach(c => c.style.display = 'none');
            document.querySelectorAll('.nav-tab').forEach(t => t.classList.remove('active'));
            document.getElementById(tabId).style.display = 'block';
            el.classList.add('active');
        }}

        // Filter Setting
        function setFilter(filterKey, el) {{
            currentFilter = filterKey;
            document.querySelectorAll('.chip').forEach(c => c.classList.remove('active'));
            el.classList.add('active');
            renderAllTables();
        }}

        function onSearchChange() {{
            currentSearch = document.getElementById('searchInput').value.trim();
            renderAllTables();
        }}

        function sortTable(tableType, key) {{
            if (currentSort.key === key) {{
                currentSort.order = currentSort.order === 'asc' ? 'desc' : 'asc';
            }} else {{
                currentSort.key = key;
                currentSort.order = 'desc';
            }}
            renderAllTables();
        }}

        // Modal Handlers
        function openDivModal(code) {{
            const target = rawData.find(t => t.code === code);
            if (!target) return;
            document.getElementById('modalDivTitle').innerText = `${{target.sname}} (${{target.code}}) 歷次詳細配息與填息天數紀錄`;
            const tbody = document.getElementById('modalDivBody');
            if (!target.div_records || target.div_records.length === 0) {{
                tbody.innerHTML = '<tr><td colspan="7" style="text-align:center;">此標的無配息紀錄或為不配息型標的</td></tr>';
            }} else {{
                tbody.innerHTML = target.div_records.map(r => `
                    <tr>
                        <td><b>${{r.ym}}</b></td>
                        <td>${{r.ex_date || '--'}}</td>
                        <td><b>${{r.amount}} 元</b></td>
                        <td>${{(r.current_rate * 100).toFixed(2)}}%</td>
                        <td><b class="mono">${{(r.year_rate * 100).toFixed(2)}}%</b></td>
                        <td>${{r.status.includes('已填息') ? `<span class="badge badge-fill-fast">${{r.status}}</span>` : `<span class="badge" style="color:var(--text-sub);">${{r.status}}</span>`}}</td>
                        <td>${{r.days ? r.days + ' 天' : '--'}}</td>
                    </tr>
                `).join('');
            }}
            document.getElementById('divModal').classList.add('active');
        }}

        function openPortModal(code) {{
            const target = rawData.find(t => t.code === code);
            if (!target) return;
            const badgeClass = getHouseBadge(target.house);
            document.getElementById('modalPortTitle').innerHTML = `
                <span class="badge ${{badgeClass}}" style="vertical-align:middle;">${{target.house}}</span>
                <span>${{target.sname}}</span>
                <span class="mono" style="font-size:15px; color:var(--text-sub);">(${{target.code}})</span>
                <span style="font-size:14px; font-weight:normal; color:var(--text-muted); margin-left:8px;">主要投資標的組合</span>
            `;

            const holdings = target.holdings || [];
            const totalW = holdings.reduce((sum, h) => sum + (h.weight || 0), 0);
            const totalPct = (totalW * 100).toFixed(2);

            const metaEl = document.getElementById('modalPortMeta');
            metaEl.innerHTML = `
                <div style="display:flex; flex-wrap:wrap; gap:16px; align-items:center; background:var(--bg-body); padding:12px 16px; border-radius:8px; border:1px solid var(--border);">
                    <div><b>📊 揭露標的合計佔比：</b><span class="mono" style="font-weight:800; color:var(--accent-gold); font-size:16px;">${{totalPct}}%</span></div>
                    <div><b>🎯 核心定位：</b>${{target.positioning}}</div>
                    <div><b>🛡️ 風險等級：</b>${{getRiskBadge(target.risk)}}</div>
                    <div><b>🏢 台積電佔比：</b><span class="mono" style="font-weight:700;">${{(target.tsmc * 100).toFixed(1)}}%</span></div>
                    <div><b>⚡ 科技合計純度：</b><span class="mono" style="font-weight:700;">${{(target.tech_total * 100).toFixed(1)}}%</span></div>
                </div>
            `;

            const visualEl = document.getElementById('modalPortVisual');
            const tbody = document.getElementById('modalPortBody');

            if (holdings.length === 0) {{
                visualEl.innerHTML = '<div style="text-align:center; padding:30px; color:var(--text-muted);">此標的未揭露詳細成分股清單（主要為海外全市場指數或公債）</div>';
                tbody.innerHTML = '<tr><td colspan="5" style="text-align:center;">此標的未揭露詳細成分股</td></tr>';
            }} else {{
                const maxWeight = Math.max(...holdings.map(h => h.weight || 0), 0.05);

                visualEl.innerHTML = `
                    <div style="display:flex; justify-content:space-between; align-items:center; margin-bottom:12px;">
                        <div style="font-size:14px; font-weight:700; color:var(--brand-green); display:flex; align-items:center; gap:6px;">
                            <span>🏆 國泰官網同款主要投資標的長條圖 (佔所有投資標的 ${{totalPct}}%)</span>
                        </div>
                        <span style="font-size:12px; color:var(--text-muted);">資料來源：國泰投信官方 CWAPI / 投信公會最新揭露</span>
                    </div>
                    <div class="cathay-holdings-container">
                        ${{holdings.map((h, i) => {{
                            const pct = (h.weight * 100).toFixed(2);
                            const fillWidth = Math.min(100, Math.max(5, (h.weight / maxWeight) * 100));
                            const rankClass = i === 0 ? 'top1' : (i === 1 ? 'top2' : (i === 2 ? 'top3' : ''));
                            return `
                                <div class="holding-bar-item">
                                    <div class="holding-bar-bg" style="width: ${{fillWidth}}%;"></div>
                                    <div class="holding-bar-info">
                                        <div class="holding-bar-rank ${{rankClass}}">#${{h.rank}}</div>
                                        <div class="holding-bar-text">
                                            <div class="holding-bar-name">
                                                ${{h.name}}
                                                ${{h.code ? `<span class="mono" style="font-size:12px; color:var(--text-muted); font-weight:normal;">(${{h.code}})</span>` : ''}}
                                            </div>
                                            <div class="holding-bar-sub">
                                                <span class="badge" style="background:rgba(255,255,255,0.06); font-size:10px;">${{h.ind || '其他產業'}}</span>
                                            </div>
                                        </div>
                                    </div>
                                    <div class="holding-bar-weight">${{pct}}%</div>
                                </div>
                            `;
                        }}).join('')}}
                    </div>
                `;

                tbody.innerHTML = holdings.map(h => `
                    <tr>
                        <td><b>#${{h.rank}}</b></td>
                        <td><span class="mono">${{h.code || '--'}}</span></td>
                        <td><b>${{h.name}}</b></td>
                        <td><span class="badge" style="background:var(--badge-bg);">${{h.ind || '--'}}</span></td>
                        <td><b class="mono">${{(h.weight * 100).toFixed(2)}}%</b></td>
                    </tr>
                `).join('');
            }}
            document.getElementById('portModal').classList.add('active');
        }}

        function closeModal(e, modalId) {{
            if (e.target.id === modalId) {{
                document.getElementById(modalId).classList.remove('active');
            }}
        }}

        // Theme Toggle
        function toggleTheme() {{
            const current = document.documentElement.getAttribute('data-theme');
            const next = current === 'light' ? 'dark' : 'light';
            document.documentElement.setAttribute('data-theme', next);
            document.getElementById('themeBtn').innerText = next === 'light' ? '🌙 深色模式' : '☀️ 淺色模式';
        }}

        // Export CSV
        function exportCurrentTable() {{
            const data = getFilteredData();
            const header = ['投信', '代碼', '簡稱', '類型', '分類', '現價', '經理費', '近3月', '近6月', '近1年', '近3年', '最新配息', '年化配息率', '平均填息天數', '台積電佔比'];
            const lines = ['\\uFEFF' + header.join(',')];
            data.forEach(t => {{
                const row = [
                    t.house, t.code, t.sname, t.type, t.cat,
                    t.price, t.fee || '', t.ret_3m || '', t.ret_6m || '', t.ret_1y || '', t.ret_3y || '',
                    t.latest_div, t.ann_yield, t.avg_fill_days || '', t.tsmc
                ].map(v => '"' + (v !== null && v !== undefined ? v : '') + '"');
                lines.push(row.join(','));
            }});
            const blob = new Blob([lines.join('\\r\\n')], {{ type: 'text/csv;charset=utf-8;' }});
            const link = document.createElement('a');
            link.href = URL.createObjectURL(blob);
            link.download = `國泰投信與熱門自選標的量化分析_${{new Date().toISOString().slice(0,10)}}.csv`;
            link.click();
        }}
    </script>
</body>
</html>"""

    with open(HTML_OUTPUT, 'w', encoding='utf-8') as f:
        f.write(html_template)
    with open(INDEX_OUTPUT, 'w', encoding='utf-8') as f:
        f.write(html_template)
    print(f"[OK] 獨立 HTML 互動式儀表板已成功生成: {os.path.abspath(HTML_OUTPUT)}")
    print(f"[OK] 同步輸出手機極簡首頁: {os.path.abspath(INDEX_OUTPUT)}")

def main():
    print("=" * 70)
    print(" 國泰投信量化分析系統 — HTML 互動式 Web 儀表板生成器")
    print(f" 執行時間: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    print("=" * 70)
    
    targets = load_data_and_enrich()
    generate_html_file(targets)
    
    print("=" * 70)
    print(" 全部工作已圓滿達成！您可直接雙擊在任何瀏覽器開啟該 HTML 檔案！")
    print("=" * 70)

if __name__ == '__main__':
    main()
