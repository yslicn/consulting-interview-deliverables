"""Check ledger review coverage and artifact bindings, not semantic truth."""
import argparse
import json
from pathlib import Path
from check_minutes_review import digest

DIMENSIONS = ('whole_record', 'source_and_time', 'verification', 'relations_and_counting')

def verify(review, data, ledger, minutes=None):
    errors = []
    if review.get('rules_version') != 'ledger-quality-v1':
        errors.append('台账评审规则须为ledger-quality-v1')
    if not review.get('reviewer') or not review.get('reviewed_at') or review.get('independent') is not True:
        errors.append('缺少独立评审身份或时间声明')
    bindings = [('data_sha256', data), ('ledger_sha256', ledger)]
    if minutes:
        bindings.append(('minutes_sha256', minutes))
    elif review.get('minutes_sha256'):
        errors.append('评审包含纪要绑定，须提供对应Word')
    for key, path in bindings:
        if review.get(key) != digest(path):
            errors.append(key + '与当前文件不一致')
    state = json.loads(Path(data).read_text(encoding='utf-8'))
    ids = [r['id'] for r in state['records']]
    checks = review.get('record_checks', [])
    checked = [r.get('id') for r in checks]
    if len(ids) != len(set(ids)) or len(checked) != len(set(checked)) or set(ids) != set(checked):
        errors.append('逐条评审编号缺失、重复或与当前台账不一致')
    for check in checks:
        if check.get('verdict') != 'PASS':
            errors.append(str(check.get('id')) + '未通过')
        for dimension in DIMENSIONS:
            item = check.get(dimension) or {}
            if not all(item.get(k) for k in ('assessment', 'basis')):
                errors.append(str(check.get('id')) + '缺少' + dimension + '判断或依据')
    joint = review.get('cross_document') or {}
    if minutes:
        if joint.get('verdict') != 'PASS' or not joint.get('checks'):
            errors.append('缺少纪要与台账联审结论及逐项依据')
        for item in joint.get('checks', []):
            if not all(item.get(k) for k in ('claim', 'minutes_location', 'record_ids', 'assessment', 'basis')):
                errors.append('联审缺少事实、位置、台账编号、判断或依据')
    elif joint.get('verdict') != 'NOT_APPLICABLE' or not joint.get('reason'):
        errors.append('未提供纪要时须说明联审不适用的实际范围')
    if review.get('verdict') != 'PASS' or not isinstance(review.get('issues'), list):
        errors.append('总评审未通过或缺少意见清单')
    for issue in review.get('issues', []):
        if issue.get('status') not in ('resolved', 'not_applicable') or not issue.get('resolution'):
            errors.append('存在未闭合意见或缺少处理依据')
    return errors

def main():
    p = argparse.ArgumentParser(description=__doc__)
    for name in ('review', 'data', 'ledger'):
        p.add_argument('--' + name, required=True, type=Path)
    p.add_argument('--minutes', type=Path)
    a = p.parse_args()
    errors = verify(json.loads(a.review.read_text(encoding='utf-8')), a.data, a.ledger, a.minutes)
    print(json.dumps({'errors': errors, 'binding_check_passed': not errors}, ensure_ascii=False, indent=2))
    return bool(errors)

if __name__ == '__main__':
    raise SystemExit(main())
