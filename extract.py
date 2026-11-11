"""Pull job-application records out of arbitrary JSON API responses."""
import datetime
import re

TITLE_KEYS = ('jobName', 'jobTitle', 'positionName', 'jobname', 'title', 'name')
COMPANY_KEYS = ('brandName', 'companyName', 'compName', 'comName', 'company', 'companyShortName',
                'brandShortName', 'name')
SALARY_KEYS = ('salaryDesc', 'salary', 'salaryText', 'salaryName', 'salary60', 'jobSalary')
CITY_KEYS = ('cityName', 'city', 'dq', 'jobCity', 'workCity', 'workCityName', 'jobArea', 'locationName',
             'unifyLocationDetail')
STATUS_KEYS = ('applyStautsDesc', 'statusDesc', 'statusName', 'applyStatusDesc', 'applyStatusName', 'stateDesc',
               'stateName',
               'deliverStatusDesc', 'statusText', 'applyStatus', 'status', 'feedbackStatusName', 'progressName')
TIME_KEYS = ('applyTime', 'happenTime', 'deliverTime', 'deliveryTime', 'applyDate', 'createTime', 'addTime', 'updateTime',
             'lastTime', 'time', 'date')
HR_KEYS = ('bossName', 'hrName', 'recruiterName', 'contactName', 'userName')
HR_TITLE_KEYS = ('bossTitle', 'hrTitle', 'recruiterTitle')
LINK_KEYS = ('jobURL', 'jobUrl', 'jobLink', 'positionURL', 'positionUrl', 'jobDetailUrl', 'link', 'url')
ID_KEYS = ('encryptJobId', 'jobId', 'jobNumber', 'positionNumber', 'number')

TITLE_PARENTS = ('job', 'position', 'jobInfo', 'jobDTO', 'jobVo', 'positionInfo', 'jobCard', '')
COMPANY_PARENTS = ('comp', 'company', 'companyInfo', 'compDTO', 'brand', 'brandInfo', 'compVo', 'companyVo')


def _flatten(d, prefix='', depth=0, out=None):
    """Flatten nested dicts (not lists) into {'a.b': value}."""
    if out is None:
        out = {}
    for k, v in d.items():
        key = prefix + k
        if isinstance(v, dict) and depth < 2:
            _flatten(v, key + '.', depth + 1, out)
        elif not isinstance(v, (dict, list)):
            out[key] = v
    return out


def _leaf(flat, keys, parents=('',)):
    """Find first non-empty string value for any key under any allowed parent."""
    for k in keys:
        for p in parents:
            for fk, v in flat.items():
                parts = fk.split('.')
                if parts[-1] != k:
                    continue
                parent = parts[-2] if len(parts) > 1 else ''
                if p and parent.lower() != p.lower():
                    continue
                if not p and k == 'name' and parent == '':
                    continue
                if isinstance(v, str) and v.strip():
                    return v.strip()
                if isinstance(v, (int, float)) and not isinstance(v, bool) and k in ID_KEYS:
                    return str(v)
    return ''


def _company(flat):
    v = _leaf(flat, COMPANY_KEYS[:-1])
    if v:
        return v
    return _leaf(flat, ('name', 'compName', 'brandName'), COMPANY_PARENTS)


def _title(flat):
    v = _leaf(flat, ('jobName', 'jobTitle', 'positionName', 'jobname'))
    if v:
        return v
    return _leaf(flat, ('title', 'name'), TITLE_PARENTS[:-1])


def _to_date(v):
    if isinstance(v, (int, float)) and not isinstance(v, bool):
        if v > 1e12:
            v = v / 1000.0
        if 1.3e9 < v < 2.5e9:
            return datetime.date.fromtimestamp(v)
        return None
    if isinstance(v, str):
        s = v.strip()
        if s.isdigit() and len(s) in (10, 13):
            return _to_date(int(s))
        m = re.search(r'(20\d\d)[-/.年](\d{1,2})[-/.月](\d{1,2})', s)
        if m:
            try:
                return datetime.date(int(m.group(1)), int(m.group(2)), int(m.group(3)))
            except ValueError:
                return None
        m = re.match(r'^(\d{1,2})[-/.月](\d{1,2})', s)
        if m:
            today = datetime.date.today()
            try:
                d = datetime.date(today.year, int(m.group(1)), int(m.group(2)))
            except ValueError:
                return None
            return d if d <= today else d.replace(year=today.year - 1)
        if s.startswith('今天') or re.match(r'^\d{1,2}:\d{2}$', s):
            return datetime.date.today()
        if s.startswith('昨天'):
            return datetime.date.today() - datetime.timedelta(days=1)
    return None


def _status(flat):
    for k in STATUS_KEYS:
        for fk, v in flat.items():
            if fk.split('.')[-1] == k and isinstance(v, str) and re.search(r'[\u4e00-\u9fff]', v):
                return v.strip()
    return ''


def record_from(d):
    flat = _flatten(d)
    title = _title(flat)
    company = _company(flat)
    if not title or not company or title == company:
        return None
    rec = {
        'title': title,
        'company': company,
        'salary': _leaf(flat, SALARY_KEYS),
        'city': _leaf(flat, CITY_KEYS),
        'status_text': _status(flat),
        'hr': _leaf(flat, HR_KEYS),
        'hr_title': _leaf(flat, HR_TITLE_KEYS),
        'link': '',
        'job_id': _leaf(flat, ID_KEYS),
        'date': None,
    }
    for k in TIME_KEYS:
        for fk, v in flat.items():
            if fk.split('.')[-1] == k:
                dt = _to_date(v)
                if dt:
                    rec['date'] = dt
                    break
        if rec['date']:
            break
    for k in LINK_KEYS:
        for fk, v in flat.items():
            parent = fk.split('.')[-2].lower() if '.' in fk else ''
            if parent.startswith(('comp', 'brand', 'boss', 'hr')):
                continue
            if fk.split('.')[-1] == k and isinstance(v, str) and v.startswith(('http', '//')):
                rec['link'] = ('https:' + v) if v.startswith('//') else v
                break
        if rec['link']:
            break
    return rec


def walk(obj, out=None):
    """Recursively find every dict that looks like a job record."""
    if out is None:
        out = []
    if isinstance(obj, dict):
        rec = record_from(obj)
        if rec:
            out.append(rec)
            return out
        for v in obj.values():
            walk(v, out)
    elif isinstance(obj, list):
        for v in obj:
            walk(v, out)
    return out
