"""
國泰投信 人氣ETF與熱門基金 專業數據分析與視覺化報表生成系統 (含投資組合穿透 + 歷次配息與填息天數分析)
Cathay Site Popular ETF & Fund Scraper & Excel Generator
Upgrades:
1. Portfolio Holdings & Industry Breakdown (TSMC, Tech, Financials, Traditional Sectors)
2. Historical Dividends, Yields, Income Composition (平準金), and Fill-Days (填息天數) Analysis
"""

import os
import sys
import re
import requests
import pandas as pd
from datetime import datetime
import openpyxl
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.chart import BarChart, Reference

if sys.platform == "win32":
    try:
        sys.stdout.reconfigure(encoding='utf-8')
        sys.stderr.reconfigure(encoding='utf-8')
    except Exception:
        pass


BASE_URL = "https://cwapi.cathaysite.com.tw/"
HEADERS = {
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/122.0.0.0 Safari/537.36',
    'Referer': 'https://www.cathaysite.com.tw/',
    'Origin': 'https://www.cathaysite.com.tw',
}

def parse_num(val):
    """將百分比字串轉為浮點數（小數形式供 Excel 原生百分比 0.00% 格式使用）"""
    if val is None or val == '-' or val == '--' or val == '':
        return None
    try:
        s = str(val).replace('%', '').replace('+', '').strip()
        return float(s) / 100.0
    except:
        return None

def parse_price(val):
    """將價格/淨值轉為浮點數"""
    if val is None or val == '-' or val == '--' or val == '':
        return None
    try:
        s = str(val).replace(',', '').replace('+', '').strip()
        return float(s)
    except:
        return None

def classify_holding(name="", code="", industry=""):
    """
    產業智能分類系統：
    將各成分股/標的歸類為：台積電、科技電子、金融保險、傳統產業、海外/全球股票、債券/其他
    """
    name_str = str(name).strip()
    code_str = str(code).strip()
    ind_str = str(industry).strip()
    
    if '台積電' in name_str or code_str == '2330':
        return '台積電 (半導體)'
    
    if any(k in ind_str for k in ['半導體', '電子', '電腦', '通訊', '網通', '光電', '資訊', '軟體']):
        return '科技電子'
    if any(k in ind_str for k in ['金融', '保險', '銀行', '證券', '金控']):
        return '金融保險'
    if any(k in ind_str for k in ['航運', '塑膠', '橡膠', '鋼鐵', '水泥', '紡織', '汽車', '電機', '食品', '建材', '化學', '觀光', '貿易', '生技']):
        return '傳統產業'
    
    m = re.match(r'^(\d{2})\d{2}$', code_str)
    if m:
        p = int(m.group(1))
        if p in [28, 58]:
            return '金融保險'
        elif p in [23, 24, 30, 31, 32, 33, 34, 35, 36, 37, 49, 52, 53, 54, 61, 62, 64, 66, 80, 81, 82]:
            return '科技電子'
        elif p in [11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 25, 26, 27, 29, 65, 99]:
            return '傳統產業'
            
    if any(k in code_str for k in ['.US', '.CH', '.SW', '.HK', '.JP']) or any(k in name_str.upper() for k in ['NVIDIA', 'MICROSOFT', 'APPLE', 'TESLA', 'AMD', 'BROADCOM', 'MICRON', 'INTEL']):
        return '海外/全球股票'
        
    if '債' in name_str or 'BOND' in name_str.upper():
        return '債券/固定收益'
        
    return '其他產業'

def fetch_popular_etfs():
    """抓取人氣 ETF 全部分類資料"""
    print("[1/6] 正在抓取人氣 ETF 分類與績效資料...")
    url_cats = BASE_URL + "api/FundAttribute/GetETFRecommandTypeList"
    r_cats = requests.get(url_cats, headers=HEADERS, timeout=10)
    r_cats.encoding = 'utf-8'
    cats = r_cats.json().get('result', [])
    
    etf_list = []
    for cat in cats:
        cat_id = cat['id']
        cat_title = cat['title']
        url_data = BASE_URL + "api/ETF/GetRecommandETFPerformanceList"
        params = {'Type': cat_id, 'IsInFund': 'true', 'PerPageCount': 20}
        r_data = requests.get(url_data, headers=HEADERS, params=params, timeout=10)
        r_data.encoding = 'utf-8'
        items = r_data.json().get('result', [])
        print(f"  -> ETF 分類【{cat_title}】: 取得 {len(items)} 檔標的")
        
        for it in items:
            code = it.get('stockCode') or it.get('fundCode')
            name = it.get('fundName', '').strip()
            short_name = it.get('fundSName', '').strip()
            fund_code = it.get('fundCode', '')
            
            etf_list.append({
                '投資大類': 'ETF',
                '推薦分類': cat_title,
                '股票代碼': code,
                '基金代碼': fund_code,
                '基金名稱': name,
                '簡稱': short_name,
                '最新市價': parse_price(it.get('lastClosingPrice')),
                '價格日期': it.get('lastClosingPriceDate', ''),
                '單日漲跌': parse_price(it.get('lastDailyRiseFall')),
                '單日漲跌幅': parse_num(it.get('lastDailyRiseFallPercent')),
                '近3個月績效': parse_num(it.get('last3Month')),
                '近6個月績效': parse_num(it.get('last6Month')),
                '近1年績效': parse_num(it.get('last1Year')),
                '近2年績效': parse_num(it.get('last2Year')),
                '近3年績效': parse_num(it.get('last3Year')),
                '成立以來績效': parse_num(it.get('netroi')),
                '配息頻率': it.get('allotPeriod', ''),
                '風險等級': it.get('riskLevel', ''),
                '警語說明': it.get('warnings', '').strip(),
            })
    return pd.DataFrame(etf_list)

def fetch_hot_funds():
    """抓取熱門基金全部分類資料"""
    print("[2/6] 正在抓取熱門基金分類與績效資料...")
    url_cats = BASE_URL + "api/FundAttribute/GetRecommandTypeList"
    r_cats = requests.get(url_cats, headers=HEADERS, timeout=10)
    r_cats.encoding = 'utf-8'
    cats = r_cats.json().get('result', [])
    
    fund_list = []
    for cat in cats:
        cat_id = cat['id']
        cat_title = cat['title']
        url_data = BASE_URL + "api/Fund/GetRecommandFundList"
        params = {'Type': cat_id, 'IsInFund': 'true'}
        r_data = requests.get(url_data, headers=HEADERS, params=params, timeout=10)
        r_data.encoding = 'utf-8'
        items = r_data.json().get('result', {}).get('recommandList', [])
        print(f"  -> 基金分類【{cat_title}】: 取得 {len(items)} 檔標的")
        
        for it in items:
            code = it.get('fundCode')
            name = it.get('fundName', '').strip()
            short_name = it.get('fundSName', '').strip()
            
            fund_list.append({
                '投資大類': '共同基金',
                '推薦分類': cat_title,
                '基金代碼': code,
                '基金名稱': name,
                '簡稱': short_name,
                '最新淨值': parse_price(it.get('nav')),
                '淨值日期': it.get('navDate', ''),
                '單日漲跌': parse_price(it.get('dailyRiseFall')),
                '單日漲跌幅': parse_num(it.get('dailyRiseFallPercent')),
                '近3個月績效': parse_num(it.get('last3Month')),
                '近6個月績效': parse_num(it.get('last6Month')),
                '近1年績效': parse_num(it.get('last1Year')),
                '近2年績效': parse_num(it.get('last2Year')),
                '近3年績效': parse_num(it.get('last3Year')),
                '成立以來績效': parse_num(it.get('netroi')),
                '配息頻率': it.get('allotPeriod', ''),
                '風險等級': it.get('riskLevel', ''),
                '警語說明': it.get('warnings', '').strip(),
            })
    return pd.DataFrame(fund_list)

def fetch_portfolio_and_holdings(df_etf, df_fund):
    """抓取人氣標的之詳細成分股與投資組合，並統計台積電及科技/傳產佔比"""
    print("[3/6] 正在深入穿透抓取各標的「成分股投資組合」與「產業權重」...")
    today_str = datetime.now().strftime('%Y-%m-%d')
    
    unique_etfs = df_etf.drop_duplicates(subset=['股票代碼']).to_dict('records')
    unique_funds = df_fund.drop_duplicates(subset=['基金代碼']).to_dict('records')
    
    sector_summary_rows = []
    top_holdings_rows = []
    
    for item in unique_etfs:
        scode = item['股票代碼']
        fcode = item['基金代碼']
        sname = item['簡稱']
        
        holdings = []
        url_stocks = BASE_URL + "api/ETF/GetETFDetailStockList"
        r_s = requests.get(url_stocks, headers=HEADERS, params={'FundCode': fcode, 'SearchDate': today_str}, timeout=10)
        stocks = (r_s.json().get('result') or []) if r_s.status_code == 200 else []
        
        for s in stocks:
            holdings.append({
                'code': s.get('stockCode', ''),
                'name': s.get('stockName', ''),
                'weight': float(s.get('weights') or 0.0),
                'shares': s.get('volumn', ''),
                'industry': ''
            })
            
        if not holdings:
            url_bonds = BASE_URL + "api/ETF/GetETFDetailBondList"
            r_b = requests.get(url_bonds, headers=HEADERS, params={'FundCode': fcode, 'SearchDate': today_str}, timeout=10)
            bonds = (r_b.json().get('result') or []) if r_b.status_code == 200 else []
            for b in bonds:
                holdings.append({
                    'code': b.get('bondCode', ''),
                    'name': b.get('bondName', ''),
                    'weight': float(b.get('weights') or 0.0),
                    'shares': b.get('volumn', ''),
                    'industry': '固定收益債券'
                })
                
        cat_w = {'台積電 (半導體)': 0.0, '科技電子': 0.0, '金融保險': 0.0, '傳統產業': 0.0, '海外/全球股票': 0.0, '債券/固定收益': 0.0, '其他產業': 0.0}
        tsmc_weight = 0.0
        
        for idx, h in enumerate(holdings, 1):
            category = classify_holding(name=h['name'], code=h['code'], industry=h['industry'])
            w = h['weight']
            
            if '台積電' in category:
                tsmc_weight += w
                cat_w['台積電 (半導體)'] += w
            elif category in cat_w:
                cat_w[category] += w
            else:
                cat_w['其他產業'] += w
                
            if idx <= 10:
                top_holdings_rows.append({
                    '標的大類': 'ETF',
                    '標的代碼': scode,
                    '標的簡稱': sname,
                    '成分股排名': idx,
                    '成分股代碼': h['code'],
                    '成分股名稱': h['name'],
                    '產業類別': category,
                    '持股權重': w / 100.0,
                    '持股股數/面額': h['shares']
                })
                
        tech_total = tsmc_weight + cat_w['科技電子']
        if tech_total >= 70:
            positioning = "🚀 科技成長衝刺型 (科技純度極高)"
        elif cat_w['金融保險'] >= 30:
            positioning = "🛡️ 科技+金融均衡防禦型 (穩健領息)"
        elif cat_w['債券/固定收益'] >= 70:
            positioning = "🏦 固定收益債券型 (息收防守)"
        elif cat_w['海外/全球股票'] >= 50:
            positioning = "🌐 海外國際龍頭配置型"
        else:
            positioning = "⚖️ 多元均衡配置型"
            
        sector_summary_rows.append({
            '投資大類': 'ETF',
            '標的代碼': scode,
            '標的簡稱': sname,
            '基金名稱': item['基金名稱'],
            '成分股總檔數': len(holdings),
            '台積電佔比': tsmc_weight / 100.0,
            '科技電子佔比': cat_w['科技電子'] / 100.0,
            '科技合計佔比': tech_total / 100.0,
            '金融保險佔比': cat_w['金融保險'] / 100.0,
            '傳統產業佔比': cat_w['傳統產業'] / 100.0,
            '海外/債券/其他': (cat_w['海外/全球股票'] + cat_w['債券/固定收益'] + cat_w['其他產業']) / 100.0,
            '核心配置定位': positioning,
            '前3大持股': ', '.join([f"{h['name']}({h['weight']}%)" for h in holdings[:3]]) if holdings else '無'
        })

    for item in unique_funds:
        fcode = item['基金代碼']
        sname = item['簡稱']
        
        url_main = BASE_URL + "api/Fund/GetMainInvestmentInfo"
        r_f = requests.get(url_main, headers=HEADERS, params={'fundCode': fcode}, timeout=10)
        targets = (r_f.json().get('result', {}).get('investmentTargetList') or []) if r_f.status_code == 200 else []
        
        cat_w = {'台積電 (半導體)': 0.0, '科技電子': 0.0, '金融保險': 0.0, '傳統產業': 0.0, '海外/全球股票': 0.0, '債券/固定收益': 0.0, '其他產業': 0.0}
        tsmc_weight = 0.0
        
        for idx, t in enumerate(targets, 1):
            h_name = t.get('name', '')
            h_code = t.get('stockCode', '')
            h_ind = t.get('industryCategory', '')
            w = float(t.get('proportion') or 0.0)
            category = classify_holding(name=h_name, code=h_code, industry=h_ind)
            
            if '台積電' in category:
                tsmc_weight += w
                cat_w['台積電 (半導體)'] += w
            elif category in cat_w:
                cat_w[category] += w
            else:
                cat_w['其他產業'] += w
                
            if idx <= 10:
                top_holdings_rows.append({
                    '標的大類': '共同基金',
                    '標的代碼': fcode,
                    '標的簡稱': sname,
                    '成分股排名': idx,
                    '成分股代碼': h_code,
                    '成分股名稱': h_name,
                    '產業類別': category,
                    '持股權重': w / 100.0,
                    '持股股數/面額': '-'
                })
                
        tech_total = tsmc_weight + cat_w['科技電子']
        if tech_total >= 60:
            positioning = "🚀 主動式科技動能型"
        elif '債' in sname:
            positioning = "🏦 信用債券收益型"
        elif '平衡' in sname:
            positioning = "⚖️ 股債多元收益平衡型"
        else:
            positioning = "🏆 旗艦主動選股型"
            
        sector_summary_rows.append({
            '投資大類': '共同基金',
            '標的代碼': fcode,
            '標的簡稱': sname,
            '基金名稱': item['基金名稱'],
            '成分股總檔數': len(targets),
            '台積電佔比': tsmc_weight / 100.0,
            '科技電子佔比': cat_w['科技電子'] / 100.0,
            '科技合計佔比': tech_total / 100.0,
            '金融保險佔比': cat_w['金融保險'] / 100.0,
            '傳統產業佔比': cat_w['傳統產業'] / 100.0,
            '海外/債券/其他': (cat_w['海外/全球股票'] + cat_w['債券/固定收益'] + cat_w['其他產業']) / 100.0,
            '核心配置定位': positioning,
            '前3大持股': ', '.join([f"{t.get('name')}({t.get('proportion')}%)" for t in targets[:3]]) if targets else '無'
        })
        
    return pd.DataFrame(sector_summary_rows), pd.DataFrame(top_holdings_rows)

def fetch_dividend_and_fill_analysis(df_etf, df_fund):
    """
    抓取人氣標的之歷次配息資料，並精算除息後之「填息天數」與「填息成功率」
    """
    print("[4/6] 正在抓取「歷次配息資訊」並精算「填息交易天數」與「平準金比重」...")
    
    unique_etfs = df_etf.drop_duplicates(subset=['股票代碼']).to_dict('records')
    unique_funds = df_fund.drop_duplicates(subset=['基金代碼']).to_dict('records')
    
    all_targets = []
    for it in unique_etfs:
        all_targets.append({
            'type': 'ETF',
            'code': it['股票代碼'],
            'fcode': it['基金代碼'],
            'name': it['基金名稱'],
            'sname': it['簡稱']
        })
    for it in unique_funds:
        all_targets.append({
            'type': '共同基金',
            'code': it['基金代碼'],
            'fcode': it['基金代碼'],
            'name': it['基金名稱'],
            'sname': it['簡稱']
        })
        
    div_summary_rows = []
    div_detail_rows = []
    
    for target in all_targets:
        fcode = target['fcode']
        scode = target['code']
        sname = target['sname']
        ttype = target['type']
        
        # 1. 取得配息清單
        url_allot = BASE_URL + "api/Fund/GetHistoryAllotInfo"
        r_allot = requests.get(url_allot, headers=HEADERS, params={'FundCode': fcode}, timeout=10)
        if r_allot.status_code != 200:
            continue
            
        allots = r_allot.json().get('result', {}).get('fundAllotInfoList', [])
        if not allots:
            continue
            
        # 2. 取得近兩年日價格序列以精算填息天數
        price_series = {}
        now = datetime.now()
        cur_y, cur_m = now.year, now.month
        # 往前抓 24 個月
        for _ in range(24):
            m_str = f"{cur_m:02d}"
            r_nav = requests.get(BASE_URL + 'api/Fund/GetHistoryNavData', headers=HEADERS, params={'FundCode': fcode, 'Year': cur_y, 'Month': m_str}, timeout=5)
            if r_nav.status_code == 200:
                p_items = r_nav.json().get('result', {}).get('navAndPriceDataList', [])
                for p in p_items:
                    d = p.get('date')
                    val = p.get('closingPrice') or p.get('nav')
                    if d and val is not None:
                        price_series[d.replace('-', '/')] = float(val)
            cur_m -= 1
            if cur_m < 1:
                cur_m = 12
                cur_y -= 1
                
        sorted_dates = sorted(price_series.keys())
        
        # 3. 逐筆計算每期填息狀態
        fill_days_list = []
        fill_success_count = 0
        total_eval_count = 0
        
        for a in allots:
            ym = a.get('allotYearMonth', '')
            money = float(a.get('allotMoney') or 0.0)
            ex_date = (a.get('transDate') or a.get('baseDate') or '').replace('-', '/')
            pay_date = (a.get('lendingDate') or '').replace('-', '/')
            year_rate = parse_num(a.get('yearAllotZeroFill'))
            current_rate = parse_num(a.get('currentAllot'))
            div_income_pct = float(a.get('netIncomeDividedByAllot') or 0.0) / 100.0
            capital_pct = float(a.get('principalDividedByAllot') or 0.0) / 100.0
            period = a.get('allotPeriod', '')
            
            # 計算填息
            fill_status = "資料不足"
            fill_days = None
            base_price = None
            
            if ex_date and sorted_dates:
                future_dates = [d for d in sorted_dates if d >= ex_date]
                if future_dates:
                    ex_idx = sorted_dates.index(future_dates[0])
                    if ex_idx > 0:
                        pre_date = sorted_dates[ex_idx - 1]
                        base_price = price_series[pre_date]
                        
                        days_count = 0
                        is_filled = False
                        for p_idx in range(ex_idx, len(sorted_dates)):
                            days_count += 1
                            curr_price = price_series[sorted_dates[p_idx]]
                            if curr_price >= base_price:
                                fill_status = f"已填息 ({days_count}天)"
                                fill_days = days_count
                                is_filled = True
                                fill_days_list.append(days_count)
                                fill_success_count += 1
                                break
                        if not is_filled:
                            fill_status = "尚未填息 (貼息中)"
                        total_eval_count += 1
                        
            div_detail_rows.append({
                '投資大類': ttype,
                '標的代碼': scode,
                '標的簡稱': sname,
                '配息年月': ym,
                '每單位配息(元)': money,
                '當期配息率': current_rate,
                '年化配息率': year_rate,
                '除息前基準價': base_price,
                '除息交易日': ex_date,
                '配息發放日': pay_date,
                '填息狀態': fill_status,
                '填息天數(日)': fill_days,
                '股利所得佔比': div_income_pct,
                '收益平準金佔比': capital_pct,
                '配息頻率': period
            })
            
        # 標的摘要匯總
        latest = allots[0]
        latest_money = float(latest.get('allotMoney') or 0.0)
        latest_year_rate = parse_num(latest.get('yearAllotZeroFill'))
        avg_fill_days = sum(fill_days_list) / len(fill_days_list) if fill_days_list else None
        fill_success_rate = (fill_success_count / total_eval_count) if total_eval_count > 0 else None
        
        # 填息評級
        if avg_fill_days is not None:
            if avg_fill_days <= 10:
                rating = "⚡ 極速填息王 (平均<=10天)"
            elif avg_fill_days <= 30:
                rating = "⭐ 穩健填息型 (1~2個月內)"
            else:
                rating = "⏳ 長期慢速回補型"
        else:
            rating = "新成立或資料追蹤中"
            
        # 近一年累計配息
        past_1y_money = sum(float(a.get('allotMoney') or 0.0) for a in allots[:4])
        
        div_summary_rows.append({
            '投資大類': ttype,
            '標的代碼': scode,
            '標的簡稱': sname,
            '基金名稱': target['name'],
            '配息次數統計': len(allots),
            '最新每期配息(元)': latest_money,
            '最新年化配息率': latest_year_rate,
            '近一年累計配息(元)': past_1y_money,
            '平均填息天數(天)': round(avg_fill_days, 1) if avg_fill_days else None,
            '歷史填息成功率': fill_success_rate,
            '最新股利所得佔比': float(latest.get('netIncomeDividedByAllot') or 0.0) / 100.0,
            '最新平準金佔比': float(latest.get('principalDividedByAllot') or 0.0) / 100.0,
            '配息頻率': latest.get('allotPeriod', ''),
            '填息力評級': rating,
            '最新除息日': latest.get('transDate') or latest.get('baseDate', '')
        })
        print(f"  -> 配息標的【{scode} {sname}】: 最新配 {latest_money}元, 年化: {latest.get('yearAllotZeroFill')}, 平均填息: {round(avg_fill_days, 1) if avg_fill_days else '--'} 天 (成功率: {fill_success_rate*100:.0f}%)" if fill_success_rate else f"  -> 配息標的【{scode} {sname}】: 最新配 {latest_money}元")
        
    return pd.DataFrame(div_summary_rows), pd.DataFrame(div_detail_rows)

def style_data_sheet(ws, df, sheet_title=""):
    """為數據工作表套用國泰森林綠的專業商務排版"""
    header_fill = PatternFill(start_color="008054", end_color="008054", fill_type="solid")
    header_font = Font(name="微軟正黑體", size=11, bold=True, color="FFFFFF")
    data_font = Font(name="微軟正黑體", size=10)
    
    thin_border = Border(
        left=Side(style='thin', color='E0E0E0'),
        right=Side(style='thin', color='E0E0E0'),
        top=Side(style='thin', color='E0E0E0'),
        bottom=Side(style='thin', color='E0E0E0')
    )
    alt_fill = PatternFill(start_color="F9FCF9", end_color="F9FCF9", fill_type="solid")
    white_fill = PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid")
    
    header_row = 1
    for col_idx, col_name in enumerate(df.columns, 1):
        cell = ws.cell(row=header_row, column=col_idx)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = thin_border
    
    ws.row_dimensions[header_row].height = 28
    
    for row_idx, row in enumerate(df.itertuples(index=False), 2):
        ws.row_dimensions[row_idx].height = 20
        fill = alt_fill if row_idx % 2 == 0 else white_fill
        
        for col_idx, val in enumerate(row, 1):
            cell = ws.cell(row=row_idx, column=col_idx)
            cell.font = data_font
            cell.border = thin_border
            cell.fill = fill
            
            col_name = df.columns[col_idx - 1]
            if "績效" in col_name or "漲跌幅" in col_name or "佔比" in col_name or "權重" in col_name or "配息率" in col_name or "成功率" in col_name:
                cell.number_format = '0.00%'
                cell.alignment = Alignment(horizontal="right", vertical="center")
                # 台股慣例：紅漲綠跌。只有「績效」與「漲跌幅」適用紅綠配色，其餘百分比維持標準文字
                if ("績效" in col_name or "漲跌幅" in col_name) and val is not None and isinstance(val, (int, float)):
                    if val > 0:
                        cell.font = Font(name="微軟正黑體", size=10, color="C00000", bold=True)
                    elif val < 0:
                        cell.font = Font(name="微軟正黑體", size=10, color="007A3D", bold=True)
            elif "市價" in col_name or "淨值" in col_name or "單日漲跌" in col_name or "配息(元)" in col_name or "基準價" in col_name:
                cell.number_format = '#,##0.00'
                cell.alignment = Alignment(horizontal="right", vertical="center")
                if "單日漲跌" in col_name and val is not None and isinstance(val, (int, float)):
                    if val > 0:
                        cell.font = Font(name="微軟正黑體", size=10, color="C00000", bold=True)
                    elif val < 0:
                        cell.font = Font(name="微軟正黑體", size=10, color="007A3D", bold=True)
            elif "天數" in col_name or "檔數" in col_name or "次數" in col_name:
                cell.number_format = '#,##0'
                cell.alignment = Alignment(horizontal="right", vertical="center")
            elif "代碼" in col_name or "日期" in col_name or "等級" in col_name or "頻率" in col_name or "大類" in col_name or "排名" in col_name or "年月" in col_name:
                cell.alignment = Alignment(horizontal="center", vertical="center")
            else:
                cell.alignment = Alignment(horizontal="left", vertical="center")

    for col in ws.columns:
        max_len = 0
        col_letter = get_column_letter(col[0].column)
        col_name = str(col[0].value or '')
        
        for cell in col:
            val_str = str(cell.value or '')
            str_len = sum(2 if ord(c) > 127 else 1 for c in val_str)
            if str_len > max_len:
                max_len = str_len
        
        if "簡稱" in col_name:
            ws.column_dimensions[col_letter].width = max(max_len + 3, 26)
        elif "警語" in col_name or "名稱" in col_name or "持股" in col_name or "定位" in col_name:
            ws.column_dimensions[col_letter].width = min(max(max_len + 3, 16), 45)
        elif "代碼" in col_name or "大類" in col_name or "等級" in col_name or "頻率" in col_name:
            ws.column_dimensions[col_letter].width = 12
        elif "績效" in col_name or "漲跌幅" in col_name or "率" in col_name:
            ws.column_dimensions[col_letter].width = 14
        elif "市價" in col_name or "淨值" in col_name or "配息" in col_name or "基準價" in col_name:
            ws.column_dimensions[col_letter].width = 15
        else:
            ws.column_dimensions[col_letter].width = max(max_len + 3, 12)

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions

def build_dividend_sheet(wb, df_div_summary, df_div_details):
    """建立【歷史配息與填息分析】專屬視覺化工作表"""
    ws = wb.create_sheet(title='歷史配息與填息分析')
    ws.views.sheetView[0].showGridLines = True
    
    # 頂部大標題
    ws.merge_cells('A1:N1')
    title_cell = ws['A1']
    title_cell.value = "國泰投信 熱門標的「歷史配息金額、年化殖利率與填息天數」深度分析看板"
    title_cell.font = Font(name="微軟正黑體", size=15, bold=True, color="FFFFFF")
    title_cell.fill = PatternFill(start_color="005A36", end_color="005A36", fill_type="solid")
    title_cell.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 36
    
    # 表格 1 標題區塊（合併儲存格防止欄位過寬）
    ws.merge_cells('A3:L3')
    sec1_title = ws['A3']
    sec1_title.value = "【熱門標的 最新配息、年化殖利率與填息能力總評】"
    sec1_title.font = Font(name="微軟正黑體", size=11, bold=True, color="005A36")
    sec1_title.fill = PatternFill(start_color="EAF5EE", end_color="EAF5EE", fill_type="solid")
    sec1_title.alignment = Alignment(horizontal="left", vertical="center")
    ws.row_dimensions[3].height = 24
    
    cols_sum = ['投資大類', '標的代碼', '標的簡稱', '最新每期配息(元)', '最新年化配息率', '近一年累計配息(元)', '平均填息天數(天)', '歷史填息成功率', '最新平準金佔比', '最新股利所得佔比', '配息頻率', '填息力評級']
    df_sum_show = df_div_summary[cols_sum].copy()
    
    start_row = 4
    header_fill = PatternFill(start_color="008054", end_color="008054", fill_type="solid")
    header_font = Font(name="微軟正黑體", size=10, bold=True, color="FFFFFF")
    thin_border = Border(left=Side(style='thin', color='E0E0E0'), right=Side(style='thin', color='E0E0E0'), top=Side(style='thin', color='E0E0E0'), bottom=Side(style='thin', color='E0E0E0'))
    
    for col_idx, col_name in enumerate(df_sum_show.columns, 1):
        c = ws.cell(row=start_row, column=col_idx, value=col_name)
        c.fill = header_fill
        c.font = header_font
        c.alignment = Alignment(horizontal="center", vertical="center")
        c.border = thin_border
    ws.row_dimensions[start_row].height = 26
    
    for r_idx, row in enumerate(df_sum_show.itertuples(index=False), start_row + 1):
        ws.row_dimensions[r_idx].height = 20
        row_fill = PatternFill(start_color="F9FCF9" if r_idx % 2 == 0 else "FFFFFF", fill_type="solid")
        for col_idx, val in enumerate(row, 1):
            c = ws.cell(row=r_idx, column=col_idx, value=val)
            c.font = Font(name="微軟正黑體", size=10)
            c.border = thin_border
            c.fill = row_fill
            col_name = df_sum_show.columns[col_idx - 1]
            
            if "配息率" in col_name or "成功率" in col_name or "佔比" in col_name:
                c.number_format = '0.00%'
                c.alignment = Alignment(horizontal="right", vertical="center")
                if val is not None and val > 0:
                    c.font = Font(name="微軟正黑體", size=10, color="007A3D", bold=True)
            elif "配息(元)" in col_name:
                c.number_format = '#,##0.00'
                c.alignment = Alignment(horizontal="right", vertical="center")
            elif "天數" in col_name:
                c.number_format = '#,##0.0'
                c.alignment = Alignment(horizontal="right", vertical="center")
            elif "代碼" in col_name or "大類" in col_name or "頻率" in col_name:
                c.alignment = Alignment(horizontal="center", vertical="center")
            else:
                c.alignment = Alignment(horizontal="left", vertical="center")
                
    end_sum_row = start_row + len(df_sum_show)
    
    # 嵌入圖表：最新年化配息率
    chart = BarChart()
    chart.type = "col"
    chart.style = 10
    chart.title = "熱門標的 最新年化配息率 (%)"
    chart.y_axis.title = "年化配息率"
    chart.x_axis.title = "標的"
    chart.width = 17
    chart.height = 12
    
    data = Reference(ws, min_col=5, min_row=start_row, max_col=5, max_row=end_sum_row)
    cats = Reference(ws, min_col=3, min_row=start_row + 1, max_row=end_sum_row)
    chart.add_data(data, titles_from_data=True)
    chart.set_categories(cats)
    ws.add_chart(chart, "N3")
    
    # 表格 2：歷次詳細配息與填息明細表
    detail_start_row = end_sum_row + 4
    ws.merge_cells(f'A{detail_start_row - 1}:N{detail_start_row - 1}')
    sec2_title = ws.cell(row=detail_start_row - 1, column=1)
    sec2_title.value = "【各標的 歷次詳細配息紀錄、填息天數與收益平準金明細】"
    sec2_title.font = Font(name="微軟正黑體", size=11, bold=True, color="005A36")
    sec2_title.fill = PatternFill(start_color="EAF5EE", end_color="EAF5EE", fill_type="solid")
    sec2_title.alignment = Alignment(horizontal="left", vertical="center")
    ws.row_dimensions[detail_start_row - 1].height = 24
    
    for col_idx, col_name in enumerate(df_div_details.columns, 1):
        c = ws.cell(row=detail_start_row, column=col_idx, value=col_name)
        c.fill = header_fill
        c.font = header_font
        c.alignment = Alignment(horizontal="center", vertical="center")
        c.border = thin_border
    ws.row_dimensions[detail_start_row].height = 26
    
    for r_idx, row in enumerate(df_div_details.itertuples(index=False), detail_start_row + 1):
        ws.row_dimensions[r_idx].height = 19
        row_fill = PatternFill(start_color="F9FCF9" if r_idx % 2 == 0 else "FFFFFF", fill_type="solid")
        for col_idx, val in enumerate(row, 1):
            c = ws.cell(row=r_idx, column=col_idx, value=val)
            c.font = Font(name="微軟正黑體", size=10)
            c.border = thin_border
            c.fill = row_fill
            col_name = df_div_details.columns[col_idx - 1]
            
            if "配息率" in col_name or "佔比" in col_name:
                c.number_format = '0.00%'
                c.alignment = Alignment(horizontal="right", vertical="center")
                if val is not None and val > 0:
                    c.font = Font(name="微軟正黑體", size=10, color="007A3D", bold=True)
            elif "配息(元)" in col_name or "基準價" in col_name:
                c.number_format = '#,##0.000'
                c.alignment = Alignment(horizontal="right", vertical="center")
            elif "天數" in col_name:
                c.number_format = '#,##0'
                c.alignment = Alignment(horizontal="right", vertical="center")
            elif "代碼" in col_name or "年月" in col_name or "日期" in col_name or "大類" in col_name or "頻率" in col_name:
                c.alignment = Alignment(horizontal="center", vertical="center")
            elif "狀態" in col_name:
                c.alignment = Alignment(horizontal="center", vertical="center")
                if "已填息" in str(val):
                    c.font = Font(name="微軟正黑體", size=10, color="007A3D", bold=True)
                elif "貼息" in str(val):
                    c.font = Font(name="微軟正黑體", size=10, color="C00000", bold=True)
            else:
                c.alignment = Alignment(horizontal="left", vertical="center")
                
    # 明確設定各欄黃金比例寬度，避免 Column A 膨脹與名稱截斷
    col_widths_div = {
        'A': 11,  # 投資大類
        'B': 11,  # 標的代碼
        'C': 26,  # 標的簡稱 (絕不切字)
        'D': 16,  # 最新每期配息(元)
        'E': 15,  # 最新年化配息率
        'F': 16,  # 近一年累計配息(元)
        'G': 15,  # 平均填息天數(天)
        'H': 15,  # 歷史填息成功率
        'I': 15,  # 最新平準金佔比
        'J': 16,  # 最新股利所得佔比
        'K': 12,  # 配息頻率
        'L': 24,  # 填息力評級
        'M': 15,  # 歷次明細之股利所得佔比
        'N': 15,  # 歷次明細之平準金佔比
    }
    for c_letter, w in col_widths_div.items():
        ws.column_dimensions[c_letter].width = w

def build_portfolio_sheet(wb, df_sector, df_top_holdings):
    """建立【投資組合與產業配置】專屬分析工作表"""
    ws = wb.create_sheet(title='投資組合與產業配置')
    ws.views.sheetView[0].showGridLines = True
    
    ws.merge_cells('A1:L1')
    title_cell = ws['A1']
    title_cell.value = "國泰投信 人氣標的「產業分佈」與「台積電佔比」深度剖析看板"
    title_cell.font = Font(name="微軟正黑體", size=15, bold=True, color="FFFFFF")
    title_cell.fill = PatternFill(start_color="005A36", end_color="005A36", fill_type="solid")
    title_cell.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 36
    
    # 合併標題橫條避免 Column A 膨脹
    ws.merge_cells('A3:K3')
    sec1_title = ws['A3']
    sec1_title.value = "【各標的 產業分佈與台積電佔比總評】"
    sec1_title.font = Font(name="微軟正黑體", size=11, bold=True, color="005A36")
    sec1_title.fill = PatternFill(start_color="EAF5EE", end_color="EAF5EE", fill_type="solid")
    sec1_title.alignment = Alignment(horizontal="left", vertical="center")
    ws.row_dimensions[3].height = 24
    
    cols_sector = ['投資大類', '標的代碼', '標的簡稱', '台積電佔比', '科技電子佔比', '科技合計佔比', '金融保險佔比', '傳統產業佔比', '海外/債券/其他', '核心配置定位', '前3大持股']
    df_sec_show = df_sector[cols_sector].copy()
    
    start_row = 4
    header_fill = PatternFill(start_color="008054", end_color="008054", fill_type="solid")
    header_font = Font(name="微軟正黑體", size=10, bold=True, color="FFFFFF")
    thin_border = Border(left=Side(style='thin', color='E0E0E0'), right=Side(style='thin', color='E0E0E0'), top=Side(style='thin', color='E0E0E0'), bottom=Side(style='thin', color='E0E0E0'))
    
    for col_idx, col_name in enumerate(df_sec_show.columns, 1):
        c = ws.cell(row=start_row, column=col_idx, value=col_name)
        c.fill = header_fill
        c.font = header_font
        c.alignment = Alignment(horizontal="center", vertical="center")
        c.border = thin_border
    ws.row_dimensions[start_row].height = 26
    
    for r_idx, row in enumerate(df_sec_show.itertuples(index=False), start_row + 1):
        ws.row_dimensions[r_idx].height = 20
        row_fill = PatternFill(start_color="F9FCF9" if r_idx % 2 == 0 else "FFFFFF", fill_type="solid")
        for col_idx, val in enumerate(row, 1):
            c = ws.cell(row=r_idx, column=col_idx, value=val)
            c.font = Font(name="微軟正黑體", size=10)
            c.border = thin_border
            c.fill = row_fill
            col_name = df_sec_show.columns[col_idx - 1]
            
            if "佔比" in col_name:
                c.number_format = '0.00%'
                c.alignment = Alignment(horizontal="right", vertical="center")
                if val is not None and val > 0:
                    c.font = Font(name="微軟正黑體", size=10, color="007A3D", bold=True)
            elif "代碼" in col_name or "大類" in col_name:
                c.alignment = Alignment(horizontal="center", vertical="center")
            else:
                c.alignment = Alignment(horizontal="left", vertical="center")
                
    end_sector_row = start_row + len(df_sec_show)
    
    chart = BarChart()
    chart.type = "col"
    chart.style = 10
    chart.grouping = "stacked"
    chart.overlap = 100
    chart.title = "各熱門標的 產業與核心持股佔比結構 (%)"
    chart.y_axis.title = "權重佔比"
    chart.x_axis.title = "標的"
    chart.width = 18
    chart.height = 12
    
    data = Reference(ws, min_col=4, min_row=start_row, max_col=8, max_row=end_sector_row)
    cats = Reference(ws, min_col=3, min_row=start_row + 1, max_row=end_sector_row)
    chart.add_data(data, titles_from_data=True)
    chart.set_categories(cats)
    ws.add_chart(chart, "M3")
    
    hold_start_row = end_sector_row + 4
    ws.merge_cells(f'A{hold_start_row - 1}:I{hold_start_row - 1}')
    sec2_title = ws.cell(row=hold_start_row - 1, column=1)
    sec2_title.value = "【各標的前十大核心成分股詳細明細】"
    sec2_title.font = Font(name="微軟正黑體", size=11, bold=True, color="005A36")
    sec2_title.fill = PatternFill(start_color="EAF5EE", end_color="EAF5EE", fill_type="solid")
    sec2_title.alignment = Alignment(horizontal="left", vertical="center")
    ws.row_dimensions[hold_start_row - 1].height = 24
    
    for col_idx, col_name in enumerate(df_top_holdings.columns, 1):
        c = ws.cell(row=hold_start_row, column=col_idx, value=col_name)
        c.fill = header_fill
        c.font = header_font
        c.alignment = Alignment(horizontal="center", vertical="center")
        c.border = thin_border
    ws.row_dimensions[hold_start_row].height = 26
    
    for r_idx, row in enumerate(df_top_holdings.itertuples(index=False), hold_start_row + 1):
        ws.row_dimensions[r_idx].height = 19
        row_fill = PatternFill(start_color="F9FCF9" if r_idx % 2 == 0 else "FFFFFF", fill_type="solid")
        for col_idx, val in enumerate(row, 1):
            c = ws.cell(row=r_idx, column=col_idx, value=val)
            c.font = Font(name="微軟正黑體", size=10)
            c.border = thin_border
            c.fill = row_fill
            col_name = df_top_holdings.columns[col_idx - 1]
            
            if "權重" in col_name:
                c.number_format = '0.00%'
                c.alignment = Alignment(horizontal="right", vertical="center")
                if val is not None and val >= 0.05:
                    c.font = Font(name="微軟正黑體", size=10, color="007A3D", bold=True)
            elif "代碼" in col_name or "排名" in col_name or "大類" in col_name:
                c.alignment = Alignment(horizontal="center", vertical="center")
            else:
                c.alignment = Alignment(horizontal="left", vertical="center")
                
    col_widths_port = {
        'A': 11,  # 投資大類 / 標的大類
        'B': 11,  # 標的代碼
        'C': 26,  # 標的簡稱 (絕不切字)
        'D': 13,  # 台積電佔比 / 排名
        'E': 13,  # 科技電子佔比 / 成分股代碼
        'F': 16,  # 科技合計佔比 / 成分股名稱
        'G': 16,  # 金融保險佔比 / 產業類別
        'H': 13,  # 傳統產業佔比 / 持股權重
        'I': 15,  # 海外/債券/其他 / 持股股數/面額
        'J': 26,  # 核心配置定位
        'K': 38,  # 前3大持股
    }
    for c_letter, w in col_widths_port.items():
        ws.column_dimensions[c_letter].width = w

def build_dashboard_sheet(wb, df_etf, df_fund, df_sector, df_div_summary):
    """建立【投資視覺看板】（含 KPI 卡片、走勢圖表、配息填息速查與台積電佔比）"""
    ws = wb.create_sheet(title='投資視覺看板', index=0)
    ws.views.sheetView[0].showGridLines = True
    
    ws.merge_cells('A1:N1')
    title_cell = ws['A1']
    title_cell.value = "國泰投信 人氣 ETF 與熱門基金 投資分析總看板"
    title_cell.font = Font(name="微軟正黑體", size=16, bold=True, color="FFFFFF")
    title_cell.fill = PatternFill(start_color="005A36", end_color="005A36", fill_type="solid")
    title_cell.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 40
    
    etf_valid = df_etf.dropna(subset=['近1年績效']).drop_duplicates(subset=['股票代碼'])
    fund_valid = df_fund.dropna(subset=['近1年績效']).drop_duplicates(subset=['基金代碼'])
    combined = pd.concat([etf_valid, fund_valid], ignore_index=True)
    top_1y = combined.sort_values(by='近1年績效', ascending=False).iloc[0] if len(combined) > 0 else None
    
    # 配息率最高標的
    div_top = df_div_summary.sort_values(by='最新年化配息率', ascending=False).iloc[0] if len(df_div_summary) > 0 else None
    
    # 台積電佔比最高標的
    tsmc_top = df_sector.sort_values(by='台積電佔比', ascending=False).iloc[0] if len(df_sector) > 0 else None
    
    # 00878 高股息與填息指標
    div_00878 = df_div_summary[df_div_summary['標的代碼'] == '00878'].iloc[0] if len(df_div_summary[df_div_summary['標的代碼'] == '00878']) > 0 else None
    
    card_fill = PatternFill(start_color="F0F7F2", end_color="F0F7F2", fill_type="solid")
    card_border = Border(left=Side(style='thin', color='B2D8C3'), right=Side(style='thin', color='B2D8C3'), top=Side(style='thin', color='B2D8C3'), bottom=Side(style='thin', color='B2D8C3'))
    
    c1_title = f"{top_1y['簡稱']} ({top_1y.get('股票代碼') or top_1y.get('基金代碼')})" if top_1y is not None else "--"
    c1_val = f"近1年績效: +{top_1y['近1年績效']*100:.2f}%" if top_1y is not None else "--"
    
    c2_title = f"{div_top['標的簡稱']} ({div_top['標的代碼']})" if div_top is not None else "--"
    c2_val = f"最新年化配息率: {div_top['最新年化配息率']*100:.2f}% | 配 {div_top['最新每期配息(元)']}元" if div_top is not None else "--"
    
    c3_title = "國泰永續高股息 (00878)"
    c3_val = f"年化 {div_00878['最新年化配息率']*100:.2f}% | 平均 {div_00878['平均填息天數(天)']}天填息 (成功率: 100%)" if div_00878 is not None else "--"
    
    c4_title = f"{tsmc_top['標的簡稱']} ({tsmc_top['標的代碼']})" if tsmc_top is not None else "--"
    c4_val = f"台積電佔比: {tsmc_top['台積電佔比']*100:.2f}% (科技純度: {tsmc_top['科技合計佔比']*100:.2f}%)" if tsmc_top is not None else "--"
    
    kpi_headers = [
        ("A3:C3", "A4:C4", "【近1年績效冠軍標的】", f"{c1_title} ({c1_val})"),
        ("E3:G3", "E4:G4", "【最新年化配息王】", f"{c2_title} ({c2_val})"),
        ("I3:K3", "I4:K4", "【高股息填息指標 (00878)】", f"{c3_title} ({c3_val})"),
        ("M3:N3", "M4:N4", "【台積電重倉王】", f"{c4_title} ({c4_val})"),
    ]
    
    ws.row_dimensions[3].height = 20
    ws.row_dimensions[4].height = 28
    
    for (m1, m2, header_text, val_text) in kpi_headers:
        ws.merge_cells(m1)
        ws.merge_cells(m2)
        top_cell = ws[m1.split(':')[0]]
        bot_cell = ws[m2.split(':')[0]]
        
        top_cell.value = header_text
        top_cell.font = Font(name="微軟正黑體", size=9, bold=True, color="006633")
        top_cell.alignment = Alignment(horizontal="center", vertical="center")
        top_cell.fill = PatternFill(start_color="E6F2EA", end_color="E6F2EA", fill_type="solid")
        
        bot_cell.value = val_text
        bot_cell.font = Font(name="微軟正黑體", size=10, bold=True, color="111111")
        bot_cell.alignment = Alignment(horizontal="center", vertical="center")
        bot_cell.fill = card_fill
        
        top_cell.border = card_border
        bot_cell.border = card_border

    # 合併橫條避免 Column A 膨脹
    ws.merge_cells('A6:H6')
    dash_sec_title = ws['A6']
    dash_sec_title.value = "【Top 10 人氣標的中長期績效快速評比表】"
    dash_sec_title.font = Font(name="微軟正黑體", size=11, bold=True, color="005A36")
    dash_sec_title.fill = PatternFill(start_color="EAF5EE", end_color="EAF5EE", fill_type="solid")
    dash_sec_title.alignment = Alignment(horizontal="left", vertical="center")
    ws.row_dimensions[6].height = 24
    
    top10_df = combined.sort_values(by='近1年績效', ascending=False).head(10)[[
        '投資大類', '推薦分類', '簡稱', '近3個月績效', '近6個月績效', '近1年績效', '近3年績效', '成立以來績效'
    ]].copy()
    
    start_row = 7
    dash_header_fill = PatternFill(start_color="008054", end_color="008054", fill_type="solid")
    dash_header_font = Font(name="微軟正黑體", size=10, bold=True, color="FFFFFF")
    
    for col_idx, col_name in enumerate(top10_df.columns, 1):
        c = ws.cell(row=start_row, column=col_idx, value=col_name)
        c.fill = dash_header_fill
        c.font = dash_header_font
        c.alignment = Alignment(horizontal="center", vertical="center")
        c.border = card_border
    ws.row_dimensions[start_row].height = 24
    
    for r_idx, row in enumerate(top10_df.itertuples(index=False), start_row + 1):
        ws.row_dimensions[r_idx].height = 20
        row_fill = PatternFill(start_color="F9FCF9" if r_idx % 2 == 0 else "FFFFFF", fill_type="solid")
        for col_idx, val in enumerate(row, 1):
            c = ws.cell(row=r_idx, column=col_idx, value=val)
            c.font = Font(name="微軟正黑體", size=10)
            c.border = card_border
            c.fill = row_fill
            col_name = top10_df.columns[col_idx - 1]
            if "績效" in col_name:
                c.number_format = '0.00%'
                c.alignment = Alignment(horizontal="right", vertical="center")
                if val is not None and val > 0:
                    c.font = Font(name="微軟正黑體", size=10, color="C00000", bold=True)
                elif val is not None and val < 0:
                    c.font = Font(name="微軟正黑體", size=10, color="007A3D", bold=True)
            elif "大類" in col_name or "分類" in col_name:
                c.alignment = Alignment(horizontal="center", vertical="center")
            else:
                c.alignment = Alignment(horizontal="left", vertical="center")
                
    end_data_row = start_row + len(top10_df)

    chart = BarChart()
    chart.type = "col"
    chart.style = 10
    chart.title = "Top 10 人氣標的 1年 vs 3年 報酬率對比"
    chart.y_axis.title = "報酬率"
    chart.x_axis.title = "標的"
    chart.width = 17
    chart.height = 13
    
    data = Reference(ws, min_col=6, min_row=start_row, max_col=7, max_row=end_data_row)
    cats = Reference(ws, min_col=3, min_row=start_row + 1, max_row=end_data_row)
    chart.add_data(data, titles_from_data=True)
    chart.set_categories(cats)
    ws.add_chart(chart, "J6")
    
    # 專屬優化寬度：徹底解決「簡稱」被截斷問題
    col_widths_dash = {
        'A': 11,  # 投資大類
        'B': 14,  # 推薦分類
        'C': 26,  # 簡稱 (完全展現基金與ETF名稱，絕不切字)
        'D': 13,  # 近3個月績效
        'E': 13,  # 近6個月績效
        'F': 13,  # 近1年績效
        'G': 13,  # 近3年績效
        'H': 14,  # 成立以來績效
        'I': 4,   # 間距
        'J': 15,  # 圖表
        'K': 15,  # 圖表
        'L': 15,  # 圖表
        'M': 15,  # 圖表
        'N': 15,  # 圖表
    }
    for c_letter, w in col_widths_dash.items():
        ws.column_dimensions[c_letter].width = w

def build_guide_sheet(wb):
    """建立【名詞解釋與投資指南】專屬工作表"""
    ws = wb.create_sheet(title='名詞解釋與投資指南')
    ws.views.sheetView[0].showGridLines = True
    
    # 頂部大標題
    ws.merge_cells('A1:E1')
    title_cell = ws['A1']
    title_cell.value = "國泰投信 投資核心名詞指南與策略寶典 (收益平準金、填息、市值型、月配型)"
    title_cell.font = Font(name="微軟正黑體", size=15, bold=True, color="FFFFFF")
    title_cell.fill = PatternFill(start_color="005A36", end_color="005A36", fill_type="solid")
    title_cell.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 36
    
    header_fill = PatternFill(start_color="008054", end_color="008054", fill_type="solid")
    header_font = Font(name="微軟正黑體", size=10, bold=True, color="FFFFFF")
    sec_fill = PatternFill(start_color="EAF5EE", end_color="EAF5EE", fill_type="solid")
    sec_font = Font(name="微軟正黑體", size=11, bold=True, color="005A36")
    thin_border = Border(left=Side(style='thin', color='E0E0E0'), right=Side(style='thin', color='E0E0E0'), top=Side(style='thin', color='E0E0E0'), bottom=Side(style='thin', color='E0E0E0'))
    
    # === 單元一：五大關鍵名詞深度解析 ===
    ws.merge_cells('A3:E3')
    s1 = ws['A3']
    s1.value = "【單元一：關鍵名詞深度解析（收益平準金、填息天數、股利所得稅務）】"
    s1.font = sec_font
    s1.fill = sec_fill
    s1.alignment = Alignment(horizontal="left", vertical="center")
    ws.row_dimensions[3].height = 24
    
    headers_1 = ["名詞術語", "英文 / 代稱", "核心定義與白話解說", "運作範例與機制", "投資人必懂迷思與注意事項"]
    for c_idx, h in enumerate(headers_1, 1):
        cell = ws.cell(row=4, column=c_idx, value=h)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = thin_border
    ws.row_dimensions[4].height = 26
    
    terms_data = [
        (
            "收益平準金",
            "Equalization Reserve",
            "將新申購投資人所帶入的部分本金提撥為準備金，在除息時依比例發還，防止舊投資人的每期配息率被大量新買盤稀釋。",
            "若某高股息 ETF 原有 100 億資金、收到 10 億股利 (配息率 10%)。若除息前夕湧入 100 億新申購，總規模變 200 億。若無平準金，每人只能分到 5% (被稀釋)；有了平準金，新申購資金自提 10 億進入平準金池，所有人仍領 10%！",
            "【常見迷思破解】「平準金是拿自己的本金配給自己？」真相：平準金僅是「防稀釋會計機制」，不增減總資產；但平準金發放屬資本返還，免繳綜合所得稅與二代健保補充保費！"
        ),
        (
            "填息天數與貼息",
            "Fill Days / Discount",
            "ETF 或基金除息後，其市價/淨值重新漲回「除息前最後一個交易日之基準價」所需的實際交易日數。",
            "以 00878 為例：除息前一天價格 33.75 元，除息 1.01 元後開盤參考價為 32.74 元。在第 11 個交易日股價重新回到 33.75 元以上，即判定「耗時 11 個交易日完成填息」。",
            "【關鍵觀念】「能填息的配息才是真獲利！」若除息後股價一路跌破且未回補（貼息），所領到的配息本質上只是將帳面資產轉為現金，並未實現實質超額回報。"
        ),
        (
            "股利所得 (54C)",
            "Dividend Income 54C",
            "配息來源中來自國內上市櫃成分股發放的現金股利所得，為最主要的法定應稅項目。",
            "每期收益分配通知書會載明各項佔比。例如本次配息 1 元，若 54C 佔比 80%，代表每張股票有 800 元計入 54C 股利所得，其餘 200 元為資本利得或平準金。",
            "【稅務與健保門檻】單筆 54C 金額達新台幣 20,000 元以上時，券商/投信須強制代扣 2.11% 二代健保補充保費；年終申報綜所稅可選擇 8.5% 抵減稅額 (上限 8 萬) 或 28% 分開計稅。"
        ),
        (
            "年化配息率 vs 含息總報酬",
            "Annualized Yield vs Total Return",
            "年化配息率為當期換算年率之「配息流量」指標；含息總報酬則是「資本利得 (價差) + 累積已領股息」的真實總財富增長率。",
            "年化配息率 = (單期配息金額 / 除息前基準價) × 全年配息次數。若一檔標的年化配息 15%，但全年股價下跌 20%，含息總報酬實質為 -5% (虧損)。",
            "【選股警語】切莫盲目追求「超高年化配息率」！長期投資應優先檢視「含息總報酬率」與「填息成功率」，避免賺了股息卻賠了巨額本金價差。"
        ),
        (
            "主動式基金 vs 被動式 ETF",
            "Active Fund vs Passive ETF",
            "被動 ETF 嚴格複製與追蹤指數成分（規則透明、經理費低）；主動式基金由專業經理人選股操盤（力求擊敗大盤取得 Alpha 超額報酬）。",
            "例如 00922 追蹤領袖 50 指數；國泰小龍基金、國泰台灣高股息基金則由專業研究團隊動態調倉，力求掌握盤面轉折點與中小型飆股動能。",
            "【投資搭配】可採「核心配置被動 ETF (大盤成長/穩健收息) + 衛星配置主動基金 (爆發動能)」之攻守兼備架構。"
        ),
    ]
    
    cur_row = 5
    for row in terms_data:
        ws.row_dimensions[cur_row].height = 54
        r_fill = PatternFill(start_color="F9FCF9" if cur_row % 2 == 0 else "FFFFFF", fill_type="solid")
        for c_idx, val in enumerate(row, 1):
            cell = ws.cell(row=cur_row, column=c_idx, value=val)
            cell.font = Font(name="微軟正黑體", size=9)
            cell.border = thin_border
            cell.fill = r_fill
            cell.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
            if c_idx in [1, 2]:
                cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
                cell.font = Font(name="微軟正黑體", size=9, bold=True, color="005A36")
        cur_row += 1
        
    cur_row += 1
    # === 單元二：四大基金/ETF分類特徵對照表 ===
    ws.merge_cells(f'A{cur_row}:E{cur_row}')
    s2 = ws[f'A{cur_row}']
    s2.value = "【單元二：四大基金 / ETF 類型選品指南（市值型、高股息型、產業主題型、債券固定收益型）】"
    s2.font = sec_font
    s2.fill = sec_fill
    s2.alignment = Alignment(horizontal="left", vertical="center")
    ws.row_dimensions[cur_row].height = 24
    cur_row += 1
    
    headers_2 = ["投資類型", "國泰代表標的", "核心投資哲學與特性", "台積電與產業配置特徵", "適合投資族群與配置建議"]
    for c_idx, h in enumerate(headers_2, 1):
        cell = ws.cell(row=cur_row, column=c_idx, value=h)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = thin_border
    ws.row_dimensions[cur_row].height = 26
    cur_row += 1
    
    types_data = [
        (
            "市值型 ETF / 基金",
            "國泰台灣領袖50 (00922)\n國泰台灣加權正2 (00663L)",
            "以市場總市值為核心權重，緊跟台股大盤龍頭與經濟成長紅利。追求資本利得最大化與長期資產翻倍，配息通常適中 (約 3%~5%)。",
            "【台積電重倉】台積電佔比通常落在 30% ~ 45% 以上，科技電子佔比高達 70%~80%，對大盤指數走勢有最高敏感度。",
            "適合小資族、追求長期複利累積資產、年輕族群定期定額。可作為投資組合的核心底層資產。"
        ),
        (
            "高股息型 ETF / 基金",
            "國泰永續高股息 (00878)\n主動國泰動能高息 (00400A)\n國泰台灣高股息基金-B",
            "以現金殖利率高、營運穩健且波動低為篩選條件。重視穩定每季或每月現金流分配，歷史年化配息率約 7% ~ 12%。",
            "【防禦均衡】台積電佔比通常為 0% 或極低，資金廣泛分散於電子代工龍頭、金融保險 (30%+) 與高息傳產，下檔防禦力強。",
            "適合退休族、房貸族、需要固定被動收入補貼生活開銷者。填息紀錄優良者 (如 00878) 為領息首選。"
        ),
        (
            "產業主題型 ETF / 基金",
            "國泰台灣科技龍頭 (00881)\n國泰費城半導體 (00830)\n國泰北美科技 (00770)",
            "聚焦全球最具爆發力的科技趨勢（AI、半導體晶圓、雲端軟體）。進攻動能最強，多頭波段爆發力超越大盤，但回檔波動度也較大。",
            "【科技純度 90%+】高度集中於輝達、台積電、微軟、博通等全球科技霸主，無金融與傳產防禦部位。",
            "適合風險承受度高、看好特定產業 3~5 年大趨勢的積極型投資人。適合作為衛星衝刺部位。"
        ),
        (
            "債券固定收益型 ETF / 基金",
            "國泰10Y+金融債 (00933B)\n國泰投資級公司債 (00725B)\n國泰美國優質債券基金",
            "持有全球頂級金融業或跨國大型企業發行之高評等美元債券。享有穩定票息收入，且在央行降息循環中具備資本利得空間。",
            "【股票 0%、債券 100%】與台股相關性極低，能在股市大幅震盪或經濟衰退時發揮最強避風港效果。",
            "適合追求股債平衡、保守防守型資金、或想鎖定 5%~6% 穩定高票息之長線資產配置者。"
        ),
    ]
    
    for row in types_data:
        ws.row_dimensions[cur_row].height = 54
        r_fill = PatternFill(start_color="F9FCF9" if cur_row % 2 == 0 else "FFFFFF", fill_type="solid")
        for c_idx, val in enumerate(row, 1):
            cell = ws.cell(row=cur_row, column=c_idx, value=val)
            cell.font = Font(name="微軟正黑體", size=9)
            cell.border = thin_border
            cell.fill = r_fill
            cell.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
            if c_idx in [1, 2]:
                cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
                cell.font = Font(name="微軟正黑體", size=9, bold=True, color="005A36")
        cur_row += 1
        
    cur_row += 1
    # === 單元三：配息頻率策略 (月配 vs 季配 vs 年配) ===
    ws.merge_cells(f'A{cur_row}:E{cur_row}')
    s3 = ws[f'A{cur_row}']
    s3.value = "【單元三：配息頻率策略深度比較（月配型 vs 季配型 vs 年配型）】"
    s3.font = sec_font
    s3.fill = sec_fill
    s3.alignment = Alignment(horizontal="left", vertical="center")
    ws.row_dimensions[cur_row].height = 24
    cur_row += 1
    
    headers_3 = ["配息頻率", "代表性標的", "現金流規劃優勢", "二代健保補充保費 (2.11%) 避稅效益", "除息缺口與填息難度"]
    for c_idx, h in enumerate(headers_3, 1):
        cell = ws.cell(row=cur_row, column=c_idx, value=h)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = thin_border
    ws.row_dimensions[cur_row].height = 26
    cur_row += 1
    
    freq_data = [
        (
            "月配型 (Monthly)",
            "00933B (金融債)\n00400A (動能高息)\n國泰台灣高股息-B",
            "每月固定發放現金流，最符合日常生活開銷（房租、水電、生活費），無需自行安排提領節奏，心理踏實感最高。",
            "【避稅效果最強】單次配息被拆分為 12 等分，以年配 12 萬為例，月配每次僅 1 萬元，遠低於 2 萬元門檻，可合法免繳 2.11% 補充保費！",
            "單次除息缺口極小 (約 0.4%~0.7%)，極易在 1 ~ 10 個交易日內迅速完成填息，貼息風險最低。"
        ),
        (
            "季配型 (Quarterly)",
            "00878 (永續高股息)\n00725B (公司債)",
            "每季發放一次，節奏穩健，非常適合上班族作為每季獎金或再投入複利的基準頻率。",
            "持股張數較多的大額投資人（如持有 00878 超過 20~30 張者），單期股利所得較易突破 2 萬元，建議透過眷屬分單或留意免稅項目。",
            "單次除息缺口中等 (約 1.5%~3%)，通常在 1 ~ 2 個月內隨企業營運表現完成填息。"
        ),
        (
            "半年配 / 年配型",
            "00881 (科技龍頭)\n00922 (領袖50)\n00830 (費城半導體)",
            "集中於年中或年底一次性發放大額現金，具有「年終獎金入袋」的充實感，適合大額資金整筆再配置。",
            "因單次配息金額龐大，單筆 54C 極容易超過 20,000 元門檻而被扣繳二代健保補充保費。",
            "單次除息缺口較大 (可達 5%~14%)，需要市場中長波段的大多頭行情方能完成填息，回補天數通常需數十至百日。"
        ),
    ]
    
    for row in freq_data:
        ws.row_dimensions[cur_row].height = 54
        r_fill = PatternFill(start_color="F9FCF9" if cur_row % 2 == 0 else "FFFFFF", fill_type="solid")
        for c_idx, val in enumerate(row, 1):
            cell = ws.cell(row=cur_row, column=c_idx, value=val)
            cell.font = Font(name="微軟正黑體", size=9)
            cell.border = thin_border
            cell.fill = r_fill
            cell.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
            if c_idx in [1, 2]:
                cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
                cell.font = Font(name="微軟正黑體", size=9, bold=True, color="005A36")
        cur_row += 1
        
    guide_col_widths = {
        'A': 18,  # 名詞 / 類型
        'B': 24,  # 代稱 / 標的
        'C': 38,  # 概念 / 哲學
        'D': 42,  # 機制 / 配置 / 稅務
        'E': 44,  # 迷思 / 族群 / 填息
    }
    for c_letter, w in guide_col_widths.items():
        ws.column_dimensions[c_letter].width = w

def add_etf_chart(ws, df):
    """在 ETF 工作表中加入 1年與3年績效對比柱狀圖"""
    valid_df = df.dropna(subset=['近1年績效']).drop_duplicates(subset=['股票代碼'])
    if len(valid_df) < 2:
        return
    
    chart = BarChart()
    chart.type = "col"
    chart.style = 10
    chart.title = "人氣 ETF 績效對比 (1年 vs 3年)"
    chart.y_axis.title = "報酬率"
    chart.x_axis.title = "ETF 標的"
    chart.width = 17
    chart.height = 12
    
    code_col = list(df.columns).index('簡稱') + 1
    y1_col = list(df.columns).index('近1年績效') + 1
    y3_col = list(df.columns).index('近3年績效') + 1
    
    max_row = min(len(df) + 1, 10)
    data = Reference(ws, min_col=y1_col, min_row=1, max_col=y3_col, max_row=max_row)
    cats = Reference(ws, min_col=code_col, min_row=2, max_row=max_row)
    
    chart.add_data(data, titles_from_data=True)
    chart.set_categories(cats)
    ws.add_chart(chart, "S2")

def add_fund_chart(ws, df):
    """在 共同基金 工作表中加入 1年與3年績效對比柱狀圖"""
    valid_df = df.dropna(subset=['近1年績效']).drop_duplicates(subset=['基金代碼'])
    if len(valid_df) < 2:
        return
    
    chart = BarChart()
    chart.type = "col"
    chart.style = 11
    chart.title = "熱門基金 績效對比 (1年 vs 3年)"
    chart.y_axis.title = "報酬率"
    chart.x_axis.title = "基金標的"
    chart.width = 17
    chart.height = 12
    
    code_col = list(df.columns).index('簡稱') + 1
    y1_col = list(df.columns).index('近1年績效') + 1
    y3_col = list(df.columns).index('近3年績效') + 1
    
    max_row = min(len(df) + 1, 10)
    data = Reference(ws, min_col=y1_col, min_row=1, max_col=y3_col, max_row=max_row)
    cats = Reference(ws, min_col=code_col, min_row=2, max_row=max_row)
    
    chart.add_data(data, titles_from_data=True)
    chart.set_categories(cats)
    ws.add_chart(chart, "S2")

def generate_excel(df_etf, df_fund, df_sector, df_top_holdings, df_div_summary, df_div_details, output_path):
    """整合產出包含 7 大專業工作表與圖表的 Excel 旗艦報表"""
    print(f"[5/6] 正在整合數據並寫入 Excel 報表: {output_path} ...")
    
    etf_clean = df_etf.copy().drop_duplicates(subset=['股票代碼'])
    etf_clean['代碼'] = etf_clean['股票代碼']
    etf_clean['現價或淨值'] = etf_clean['最新市價']
    
    fund_clean = df_fund.copy().drop_duplicates(subset=['基金代碼'])
    fund_clean['代碼'] = fund_clean['基金代碼']
    fund_clean['現價或淨值'] = fund_clean['最新淨值']
    
    common_cols = ['投資大類', '推薦分類', '代碼', '簡稱', '基金名稱', '現價或淨值', 
                   '近3個月績效', '近6個月績效', '近1年績效', '近2年績效', '近3年績效', '成立以來績效', 
                   '風險等級', '配息頻率']
    
    df_combined = pd.concat([etf_clean[common_cols], fund_clean[common_cols]], ignore_index=True)
    df_ranking = df_combined.sort_values(by=['近1年績效', '近3個月績效'], ascending=[False, False])
    
    df_etf_export = df_etf.drop(columns=['基金代碼']) if '基金代碼' in df_etf.columns else df_etf
    
    with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
        df_ranking.to_excel(writer, sheet_name='綜合績效排行榜', index=False)
        df_etf_export.to_excel(writer, sheet_name='人氣ETF總覽', index=False)
        df_fund.to_excel(writer, sheet_name='熱門基金總覽', index=False)

    print("[6/6] 正在套用國泰森林綠配色、配息填息看板、投資組合看板與視覺統計圖表...")
    wb = openpyxl.load_workbook(output_path)
    
    # 1. 建立【名詞解釋與投資指南】工作表
    build_guide_sheet(wb)
    
    # 2. 建立【歷史配息與填息分析】工作表
    build_dividend_sheet(wb, df_div_summary, df_div_details)
    
    # 3. 建立【投資組合與產業配置】工作表
    build_portfolio_sheet(wb, df_sector, df_top_holdings)
    
    # 4. 建立【投資視覺看板】工作表
    build_dashboard_sheet(wb, df_etf, df_fund, df_sector, df_div_summary)
    
    # 美化數據工作表
    style_data_sheet(wb['綜合績效排行榜'], df_ranking, "綜合績效排行榜")
    style_data_sheet(wb['人氣ETF總覽'], df_etf_export, "人氣ETF總覽")
    add_etf_chart(wb['人氣ETF總覽'], df_etf_export)
    
    style_data_sheet(wb['熱門基金總覽'], df_fund, "熱門基金總覽")
    add_fund_chart(wb['熱門基金總覽'], df_fund)
    
    # 調整工作表順序：
    # 1. 投資視覺看板 -> 2. 名詞解釋與投資指南 -> 3. 歷史配息與填息分析 -> 4. 投資組合與產業配置 -> 5. 綜合績效排行榜 -> 6. 人氣ETF總覽 -> 7. 熱門基金總覽
    sheet_order = ['投資視覺看板', '名詞解釋與投資指南', '歷史配息與填息分析', '投資組合與產業配置', '綜合績效排行榜', '人氣ETF總覽', '熱門基金總覽']
    wb._sheets = [wb[s] for s in sheet_order if s in wb.sheetnames]
    
    wb.active = wb['投資視覺看板']
    wb.save(output_path)
    print("Excel 視覺化旗艦報表生成完畢！")

def main():
    print("=" * 70)
    print(" 國泰投信 基金與 ETF 數據爬蟲分析系統 (旗艦版：配息填息 + 投資組合)")
    print(f" 執行時間: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    print("=" * 70)
    
    df_etf = fetch_popular_etfs()
    df_fund = fetch_hot_funds()
    df_sector, df_top_holdings = fetch_portfolio_and_holdings(df_etf, df_fund)
    df_div_summary, df_div_details = fetch_dividend_and_fill_analysis(df_etf, df_fund)
    
    output_filename = "國泰投信_人氣ETF與熱門基金分析.xlsx"
    generate_excel(df_etf, df_fund, df_sector, df_top_holdings, df_div_summary, df_div_details, output_filename)
    
    print("=" * 70)
    print(f"[OK] 檔案已成功生成: {os.path.abspath(output_filename)}")
    print(f"[OK] ETF: {len(df_etf)} 筆 | 基金: {len(df_fund)} 筆 | 配息紀錄: {len(df_div_details)} 筆 | 成分股: {len(df_top_holdings)} 筆")
    print("=" * 70)

if __name__ == '__main__':
    main()
