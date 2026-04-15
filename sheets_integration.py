import requests
import re

SHEET_ID = '1LTz07x3GbZsVnqfKk4zP_3uq4nb1ey2Qi94nunyGKVM'

def fetch_sheet(sheet_name):
    url = f'https://docs.google.com/spreadsheets/d/{SHEET_ID}/gviz/tq?tqx=out:csv&sheet={requests.utils.quote(sheet_name)}'
    try:
        resp = requests.get(url, timeout=15)
        resp.encoding = 'utf-8'
        lines = resp.text.strip().split('\n')
        rows = []
        for line in lines:
            in_quote = False
            current = ''
            row = []
            for ch in line:
                if ch == '"':
                    in_quote = not in_quote
                elif ch == ',' and not in_quote:
                    row.append(current.strip().strip('"'))
                    current = ''
                else:
                    current += ch
            row.append(current.strip().strip('"'))
            rows.append(row)
        return rows
    except Exception as e:
        print(f'시트 로드 실패 {sheet_name}: {e}')
        return []

def parse_date_from_cell(cell_text):
    """F열에서 날짜 추출. 예: '1/2(금) 9시 화란플라워' -> '01-02' """
    match = re.search(r'(\d{1,2})/(\d{1,2})', cell_text)
    if match:
        month = match.group(1).zfill(2)
        day = match.group(2).zfill(2)
        return f'{month}-{day}'
    return None

def parse_depart_from_cell(cell_text):
    """F열에서 출발지 상호명 추출. // 로 구분된 두번째 항목"""
    # 패턴: 날짜시간// 상호명// 주소
    parts = cell_text.split('//')
    if len(parts) >= 2:
        # 두번째 파트가 상호명
        name = parts[1].strip()
        # 주소 부분 제거 (숫자나 구/로/길 로 시작하면 주소)
        name = re.split(r'[0-9]|서초구|강남구|중구|종로구|마포구|용산구|성동구|광진구|동대문구|중랑구|성북구|강북구|도봉구|노원구|은평구|서대문구|양천구|강서구|구로구|금천구|영등포구|동작구|관악구|서초구|송파구|강동구|경기|인천', name)[0].strip()
        return name if len(name) > 1 else ''
    # // 없으면 시간 다음 텍스트
    match = re.search(r'\d+시[이후]?\s*(.+)', cell_text)
    if match:
        return match.group(1).strip()
    return ''

def parse_arrive_from_cell(cell_text):
    """G열 전체가 도착지 주소"""
    if not cell_text:
        return ''
    # 전화번호 제거
    text = re.sub(r'010-\d{4}-\d{4}', '', cell_text)
    text = re.sub(r'0\d{1,2}-\d{3,4}-\d{4}', '', text)
    # 대괄호 내용 제거
    text = re.sub(r'\[.*?\]', '', text)
    return text.strip()

def build_order_db():
    """구글시트 주문리스트에서 주문 DB 구축"""
    order_db = []
    sheet_names = ['2026년2분기(4/1~)', '2026년1분기(1/2~']
    
    for sheet_name in sheet_names:
        rows = fetch_sheet(sheet_name)
        for row in rows:
            if len(row) < 5:
                continue
            
            # D열(인덱스3): 배송기사명
            driver_cell = row[3].strip() if len(row) > 3 else ''
            # F열(인덱스5): 출발지 정보
            f_cell = row[5].strip() if len(row) > 5 else ''
            # G열(인덱스6): 도착지 정보  
            g_cell = row[6].strip() if len(row) > 6 else ''
            
            if not f_cell or not driver_cell:
                continue
            
            # 배송기사명 정리 (// 이후 제거, 개인 등 제거)
            driver = re.split(r'//|//', driver_cell)[0].strip()
            driver = re.sub(r'\s*(개인|정기|퇴사|면제|파견|취소).*', '', driver).strip()
            
            if not driver or len(driver) < 2:
                continue
            
            # 날짜 추출
            date_mmdd = parse_date_from_cell(f_cell)
            if not date_mmdd:
                continue
            
            # 출발지 추출
            depart = parse_depart_from_cell(f_cell)
            
            # 도착지 추출
            arrive = parse_arrive_from_cell(g_cell)
            
            # 금액 추출 (G열에 있을 수 있음)
            price_match = re.search(r'월말\s*(\d{1,2}),?(\d{3})', g_cell)
            price = 0
            if price_match:
                price = int(price_match.group(1) + price_match.group(2))
            
            order_db.append({
                'date_mmdd': date_mmdd,
                'driver': driver,
                'depart': depart,
                'arrive': arrive,
                'price': price,
                'raw_f': f_cell,
                'raw_g': g_cell,
            })
    
    return order_db

def build_vendor_db():
    """수요처·정기배송 탭에서 수요처 DB 구축"""
    vendor_db = {}
    rows = fetch_sheet('수요처, 정기배송')
    for row in rows[1:]:
        if len(row) < 1:
            continue
        name = row[0].strip()
        address = row[1].strip() if len(row) > 1 else ''
        phone = row[2].strip() if len(row) > 2 else ''
        payment = row[3].strip() if len(row) > 3 else ''
        if name and len(name) > 1:
            vendor_db[name] = {
                'address': address,
                'phone': phone,
                'payment': payment
            }
    return vendor_db

def find_best_vendor(name_str, addr_str='', wolmal_map=None):
    """수요처 DB에서 가장 유사한 수요처 찾기"""
    if not name_str:
        return name_str
    name_clean = name_str.replace(' ', '').lower()
    
    if wolmal_map:
        for std, info in wolmal_map.items():
            for alias in info.get('aliases', []):
                if alias and alias.replace(' ', '').lower() in name_clean:
                    return std
            for key in info.get('addr', []):
                if key and addr_str and key.replace(' ', '').lower() in addr_str.replace(' ', '').lower():
                    return std
    
    try:
        vendor_db = build_vendor_db()
        for vname, info in vendor_db.items():
            v_clean = vname.replace(' ', '').lower()
            if len(v_clean) > 1 and (v_clean in name_clean or name_clean in v_clean):
                return vname
    except Exception as e:
        print(f'vendor db error: {e}')
    
    return name_str

def find_matching_order(date_str, driver_str, order_db, depart_hint='', arrive_hint=''):
    """날짜 + 배송기사 + 출발지/도착지 힌트로 주문 찾기"""
    if not date_str or not driver_str or not order_db:
        return None
    
    # 날짜에서 월-일 추출
    date_clean = re.sub(r'[^0-9]', '', str(date_str))
    if len(date_clean) >= 4:
        mmdd = date_clean[-4:]  # 마지막 4자리 = MMDD
        mm = mmdd[:2]
        dd = mmdd[2:]
        target_mmdd = f'{mm}-{dd}'
    else:
        return None
    
    driver_clean = driver_str.replace(' ', '')
    
    # 날짜 + 배송기사로 후보 찾기
    candidates = []
    for order in order_db:
        order_driver = order['driver'].replace(' ', '')
        if (order['date_mmdd'] == target_mmdd and
            (driver_clean in order_driver or order_driver in driver_clean)):
            candidates.append(order)
    
    if not candidates:
        return None
    
    # 후보가 1개면 바로 반환
    if len(candidates) == 1:
        return candidates[0]
    
    # 여러 개면 출발지/도착지 힌트로 추가 매칭
    if depart_hint or arrive_hint:
        depart_clean = depart_hint.replace(' ', '').lower()
        arrive_clean = arrive_hint.replace(' ', '').lower()
        
        best = None
        best_score = 0
        
        for c in candidates:
            score = 0
            c_depart = c['depart'].replace(' ', '').lower()
            c_arrive = c['arrive'].replace(' ', '').lower()
            
            # 출발지 매칭 점수
            if depart_clean and c_depart:
                common = sum(1 for ch in depart_clean if ch in c_depart)
                score += common
            
            # 도착지 매칭 점수
            if arrive_clean and c_arrive:
                common = sum(1 for ch in arrive_clean if ch in c_arrive)
                score += common
            
            if score > best_score:
                best_score = score
                best = c
        
        return best if best_score > 2 else candidates[0]
    
    return candidates[0]
