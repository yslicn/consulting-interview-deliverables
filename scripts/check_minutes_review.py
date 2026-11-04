"""Verify review bindings and required review evidence before minutes finalization."""
import argparse
import hashlib
import json
from pathlib import Path


def digest(path):
    return hashlib.sha256(Path(path).read_bytes()).hexdigest()


def verify(review, data, minutes):
    errors = []
    if review.get('rules_version') != 'minutes-quality-v3':
        errors.append('评审须采用当前规则 minutes-quality-v3')
    if not review.get('reviewer') or not review.get('reviewed_at') or review.get('independent') is not True:
        errors.append('缺少独立评审身份或时间声明')
    for field, path in [('data_sha256', data), ('minutes_sha256', minutes)]:
        if review.get(field) != digest(path):
            errors.append(field + '与当前稿件不一致')
    for name in ('delivery_suitability', 'evidence'):
        result = review.get(name) or {}
        if result.get('verdict') != 'PASS' or not result.get('paragraph_checks'):
            errors.append(name + '缺少PASS及逐段验收依据')
        for check in result.get('paragraph_checks', []):
            if not isinstance(check, dict) or not all(check.get(k) for k in ('location', 'assessment', 'basis')):
                errors.append(name + '的段落检查缺少位置、判断或依据')
    for name in ('synthesis', 'tracking_selection', 'challenge_checks'):
        result = review.get(name) or {}
        if result.get('verdict') != 'PASS' or not result.get('checks'):
            errors.append(name + '缺少PASS及具体检查依据')
        for check in result.get('checks', []):
            if not isinstance(check, dict) or not all(check.get(k) for k in ('location', 'assessment', 'basis')):
                errors.append(name + '缺少位置、判断或依据')
    if review.get('verdict') != 'PASS':
        errors.append('总评审未通过')
    for issue in review.get('issues', []):
        if issue.get('status') not in ('resolved', 'not_applicable') or not issue.get('resolution'):
            errors.append('存在未闭合或缺少处理依据的意见')
    return errors


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--review', required=True, type=Path)
    parser.add_argument('--data', required=True, type=Path)
    parser.add_argument('--minutes', required=True, type=Path)
    args = parser.parse_args()
    errors = verify(json.loads(args.review.read_text()), args.data, args.minutes)
    print(json.dumps({'errors': errors, 'binding_check_passed': not errors}, ensure_ascii=False, indent=2))
    return bool(errors)


if __name__ == '__main__':
    raise SystemExit(main())
