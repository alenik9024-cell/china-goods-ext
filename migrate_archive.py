# Миграция архива «Архив клиентов» в расширение (таблица через Apps Script API).
# Что переносим:
#  1) клиенты -> selections (подборы): productsText/logisticsText/customsText + ссылки на файлы Диска
#  2) файлы из archive_db.json -> прикреплены в тексте подбора ссылками (driveUrl), фото — как прямые ссылки
#  3) xlsx из «Подобранные товары» -> позиции в products (каталог), эвристический разбор колонок
import json, os, re, sys, time, zipfile
from xml.etree import ElementTree as ET
sys.path.insert(0, r"C:\Users\1\AppData\Local\hermes\cache\scratch\china-goods-ext")
import urllib.request

API_URL = "https://script.google.com/macros/s/AKfycbwfj1cxO9FcX8sUnnUIiYWD_BcU1EgJKjUQ5dsdrFhB39QgUrsKLhWIMrJRIRcKmkaZIg/exec"
TOKEN = open(r"C:\Users\1\AppData\Local\hermes\cache\scratch\china-goods-ext\api_token.txt", encoding="utf-8").read().strip() if os.path.exists(r"C:\Users\1\AppData\Local\hermes\cache\scratch\china-goods-ext\api_token.txt") else ""

ARCH = r"D:\Владимир товары из Китая\Архив клиентов-20261006T090516Z-1-001\Архив клиентов"
NS = {'m': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}

def api(action, sheet, _retry=3, **extra):
    # бэкенд ждёт полезную нагрузку в ключе "data" (см. code.gs: p.data)
    if "record" in extra:
        extra["data"] = extra.pop("record")
    body = {"token": TOKEN, "action": action, "sheet": sheet, **extra}
    last = None
    for attempt in range(_retry):
        try:
            req = urllib.request.Request(API_URL, data=json.dumps(body).encode("utf-8"),
                                         headers={"Content-Type": "text/plain"}, method="POST")
            with urllib.request.urlopen(req, timeout=45) as r:
                j = json.loads(r.read().decode("utf-8"))
            if not j.get("ok"):
                raise RuntimeError(j.get("error"))
            return j
        except Exception as e:
            last = e
            time.sleep(2 * (attempt + 1))
    raise last

def read_xlsx(path):
    """Вернёт список строк (списков значений) первого листа."""
    try:
        z = zipfile.ZipFile(path)
        ss = []
        if 'xl/sharedStrings.xml' in z.namelist():
            root = ET.fromstring(z.read('xl/sharedStrings.xml'))
            for s in root:
                ss.append(''.join(t.text or '' for t in s.iter('{http://schemas.openxmlformats.org/spreadsheetml/2006/main}t')))
        # листы по порядку
        sheets = sorted(n for n in z.namelist() if re.match(r'xl/worksheets/sheet\d+\.xml$', n))
        if not sheets: return []
        root = ET.fromstring(z.read(sheets[0]))
        rows = []
        for row in root.iter('{http://schemas.openxmlformats.org/spreadsheetml/2006/main}row'):
            vals = []
            last_col = 0
            for c in row:
                ref = c.get('r', '')
                m = re.match(r'([A-Z]+)', ref)
                col = 0
                if m:
                    for ch in m.group(1): col = col*26 + (ord(ch)-64)
                while len(vals) < col-1: vals.append('')
                v = c.find('m:v', NS)
                t = '' if v is None else (v.text or '')
                if c.get('t') == 's' and t.isdigit():
                    try: t = ss[int(t)]
                    except: pass
                vals.append(str(t).strip())
            rows.append(vals)
        return rows
    except Exception as e:
        return []

PRICE_RE = re.compile(r'\d[\d\s]*[.,]?\d*')
def parse_positions(rows, client_name, fname):
    """Эвристика: находим строку заголовка с 'цен' и 'наимен/ссылка/цвет', парсим ниже."""
    if not rows: return []
    # найти заголовок
    hdr_i, hdr = -1, []
    for i, row in enumerate(rows[:8]):
        joined = ' '.join(row).lower()
        if ('цен' in joined or 'юан' in joined or 'стоимост' in joined) and len([x for x in row if x]) >= 2:
            hdr_i, hdr = i, row
            break
    if hdr_i < 0: return []
    low = [h.lower() for h in hdr]
    def col(*keys):
        for k in keys:
            for j, h in enumerate(low):
                if k in h: return j
        return -1
    c_name = col('наимен', 'назван', 'цвет', 'модель', 'размер')
    c_price = col('цен', 'стоимость 1', 'юан')
    c_qty = col('количес', 'кол-во')
    c_link = col('ссылка')
    c_weight = col('вес')
    positions = []
    for row in rows[hdr_i+1:]:
        if not any(row): continue
        name = row[c_name] if 0 <= c_name < len(row) else ''
        price_raw = row[c_price] if 0 <= c_price < len(row) else ''
        m = PRICE_RE.search(price_raw.replace(' ', ''))
        price = 0.0
        if m:
            try: price = float(m.group(0).replace(',', '.'))
            except: pass
        link = row[c_link] if 0 <= c_link < len(row) else ''
        qty = row[c_qty] if 0 <= c_qty < len(row) else ''
        if not name and not price: continue
        if not price and not link: continue  # строки-описания пропускаем
        positions.append({
            "factory": client_name[:60] or "архив",
            "article": f"{fname[:40]}#{len(positions)+1}",
            "name": (name or fname)[:120],
            "price": price,
            "currency": "CNY",
            "weight": 0,
            "link": link[:200],
            "qty_note": str(qty)[:80],
            "price_date": "",
            "photo": "",
            "notes": f"Импорт из архива: {fname}"
        })
    return positions[:50]  # ограничение на файл

def main():
    db = json.load(open(os.path.join(ARCH, 'archive_db.json'), encoding='utf-8'))
    clients = db['clients']
    files = db['files']
    files_by_client = {}
    for f in files:
        files_by_client.setdefault(f['clientId'], []).append(f)

    stats = {"selections": 0, "products": 0, "skipped": 0, "errors": []}
    done_log = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'migrate_done.json')
    done = set()
    if os.path.exists(done_log):
        done = set(json.load(open(done_log, encoding='utf-8')))

    for i, c in enumerate(clients):
        cid = c['id']
        if cid in done: continue
        name = (c.get('name') or '').strip() or 'Без имени'
        print(f"[{i+1}/{len(clients)}] {name}…", flush=True)
        cfiles = files_by_client.get(cid, [])
        texts = []
        if c.get('productsText', '').strip():
            texts.append("ПОДОБРАННЫЕ ТОВАРЫ:\n" + c['productsText'].strip())
        if c.get('logisticsText', '').strip():
            texts.append("ЛОГИСТИКА:\n" + c['logisticsText'].strip())
        if c.get('customsText', '').strip():
            texts.append("ТАМОЖНЯ:\n" + c['customsText'].strip())
        flinks = [f"{f['originalName']} — {f['driveUrl']}" for f in cfiles]
        if flinks:
            texts.append("ФАЙЛЫ:\n" + "\n".join(flinks))
        body_text = "\n\n".join(texts).strip()

        # 1) подбор (selection) с текстом
        if body_text:
            try:
                api("add", "selections", record={
                    "deal_id": "",
                    "client_name": name,
                    "created_at": c.get('contractDate') or c.get('createdAt', ''),
                    "items_json": json.dumps([{"note": body_text[:4500]}], ensure_ascii=False),
                    "total_rub": 0,
                    "source": "archive-import"
                })
                stats["selections"] += 1
            except Exception as e:
                stats["errors"].append(f"sel {name}: {e}")

        # 2) xlsx -> каталог
        for f in cfiles:
            fn = f['originalName']
            if not fn.lower().endswith('.xlsx'): continue
            local = None
            for root, _, fs in os.walk(ARCH):
                if fn in fs and (name.split(' ')[0] in root or True):
                    cand = os.path.join(root, fn)
                    # предпочесть папку клиента
                    if name.split(' ')[0].lower() in root.lower():
                        local = cand; break
                    if local is None: local = cand
            if not local: continue
            rows = read_xlsx(local)
            positions = parse_positions(rows, name, fn)
            # фото клиента (jpg) — первую картинку ставим фото позиции
            photos = [x['driveUrl'] for x in cfiles if x['originalName'].lower().endswith(('.jpg', '.jpeg', '.png'))]
            for p in positions:
                if photos:
                    # driveUrl вида .../document/d/ID/... -> thumbnail
                    m = re.search(r'/d/([A-Za-z0-9_-]+)', photos[0])
                    if m:
                        p['photo'] = f"https://drive.google.com/thumbnail?id={m.group(1)}&sz=w400"
                try:
                    api("upsert", "products", keys=["factory", "article"], record=p)
                    stats["products"] += 1
                except Exception as e:
                    stats["errors"].append(f"prod {name}/{fn}: {e}")

        done.add(cid)
        if len(done) % 10 == 0:
            json.dump(list(done), open(done_log, 'w', encoding='utf-8'))
            print(f"[{i+1}/{len(clients)}] selections={stats['selections']} products={stats['products']} errors={len(stats['errors'])}", flush=True)

    json.dump(list(done), open(done_log, 'w', encoding='utf-8'))
    print("DONE", json.dumps({k: v for k, v in stats.items() if k != 'errors'}, ensure_ascii=False))
    print("errors:", stats["errors"][:15])

if __name__ == "__main__":
    main()
