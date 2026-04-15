import requests
import re

SHEET_ID = '1LTz07x3GbZsVnqfKk4zP_3uq4nb1ey2Qi94nunyGKVM'

def fetch_sheet(sheet_name):
    url = f'https://docs.google.com/spreadsheets/d/{SHEET_ID}/gviz/tq?tqx=out:csv&sheet={requests.utils.quote(sheet_name)}'
    try:
        resp = requests.get(url, timeout=10)
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

def build_vendor_db():
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

def build_order_db():
    order_db = []
    sheet_names = ['2026년2분기(4/1~)', '2026년1분기(1/2~']
    for sheet_name in sheet_names:
        rows = fetch_sheet(sheet_name)
        for row in rows[2:]:
            if len(row) < 4:
                continue
            date = row[1].strip() if len(row) > 1 else ''
            driver = row[3].strip() if len(row) > 3 else ''
            depart = row[5].strip() if len(row) > 5 else ''
            arrive = row[6].strip() if len(row) > 6 else ''
            if date and driver:
                order_db.append({
                    'date': date,
                    'driver': driver,
                    'depart': depart,
                    'arrive': arrive
                })
    return order_db

def find_best_vendor(name_str, addr_str='', wolmal_map=None):
    if not name_str:
        return name_str
    name_clean = name_str.replace(' ', '').lower()
    
    # WOLMAL_MAP에서 먼저 찾기
    if wolmal_map:
        for std, info in wolmal_map.items():
            for alias in info.get('aliases', []):
                if alias and alias.replace(' ', '').lower() in name_clean:
                    return std
            for key in info.get('addr', []):
                if key and addr_str and key.replace(' ', '').lower() in addr_str.replace(' ', '').lower():
                    return std
    
    # 구글시트 수요처 DB에서 찾기
    try:
        vendor_db = build_vendor_db()
        for vname, info in vendor_db.items():
            v_clean = vname.replace(' ', '').lower()
            if len(v_clean) > 1 and (v_clean in name_clean or name_clean in v_clean):
                return vname
            if addr_str and info.get('address'):
                addr_clean = addr_str.replace(' ', '').lower()
                db_addr = info['address'].replace(' ', '').lower()
                parts = [p for p in db_addr.split() if len(p) > 3]
                if parts and any(p in addr_clean for p in parts):
                    return vname
    except Exception as e:
        print(f'vendor db error: {e}')
    
    return name_str

def find_matching_order(date_str, driver_str, order_db):
    if not date_str or not driver_str or not order_db:
        return None
    date_clean = re.sub(r'[^0-9]', '', str(date_str))
    driver_clean = driver_str.replace(' ', '')
    
    for order in order_db:
        order_date = re.sub(r'[^0-9]', '', str(order['date']))
        order_driver = order['driver'].replace(' ', '')
        if (len(date_clean) >= 4 and len(order_date) >= 4 and
            date_clean[-4:] == order_date[-4:] and
            (driver_clean in order_driver or order_driver in driver_clean)):
            return order
    return None
