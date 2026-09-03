import openpyxl, re, unicodedata

def norm(s):
    if s is None: return ''
    s = unicodedata.normalize('NFKC', str(s))
    return s.replace('　',' ').strip()

TAB_RE = re.compile(r'^(?:(\d{4})年)?\s*(\d{1,2})\s*月\s*[~～\-]\s*(?:(\d{4})年)?\s*(\d{1,2})\s*月')

def parse_tab(name):
    n = norm(name)
    m = TAB_RE.match(n)
    if not m: return None
    y1, m1, y2, m2 = m.group(1), int(m.group(2)), m.group(3), int(m.group(4))
    return (int(y1) if y1 else None, m1, int(y2) if y2 else None, m2)

HDR_RE = re.compile(r'^(令和|平成|昭和)?\s*(\d{1,2}|元)\s*年\s*(\d{1,2})\s*月')
def parse_hdr(v):
    n = norm(v)
    m = HDR_RE.match(n)
    if not m: return None
    era, yr, mo = m.group(1), m.group(2), int(m.group(3))
    yr = 1 if yr == '元' else int(yr)
    if era == '令和': y = 2018 + yr
    elif era == '平成': y = 1988 + yr
    elif era == '昭和': y = 1925 + yr
    else: y = 1988 + yr   # bare number => 平成 counting (26年=2014, 32年=2020)
    return (y, mo)
