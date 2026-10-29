"""Data contract, source checks and conservative project-ledger updates."""
from copy import deepcopy
from datetime import date
import hashlib
from pathlib import Path
import re

VERSION = '1.2.0'
NATURES = {'问题', '待核实线索', '综合议题', '管理诉求', '实施风险'}
STATUSES = {'待核实', '核实中', '已确认', '已合并', '已排除'}
UTTERANCES = {'现状陈述', '管理诉求', '顾问建议', '提问', '明确决定', '明确安排', '其他'}
KINDS = {'fact', 'request', 'suggestion', 'decision'}
VERIFY_DEFAULTS = dict(focus=None, materials=None, completion_criteria=None, owner=None,
                       due_date=None, cause_hypothesis=None, priority='待评估',
                       conclusion=None, confirmer=None, confirmed_date=None)
RECORD_FIELDS = {'id', 'group', 'collaborating_groups', 'nature', 'domain', 'title',
                 'description', 'impact', 'management_request', 'departments', 'status',
                 'related_ids', 'interview_ids', 'verification', 'historical_id'}
INTERVIEW_DEFAULTS = dict(stage=None, level=None, object=None, date=None, place=None,
                          duration_hours=None, participants=[], topic=None, has_outline=False,
                          source_file=None, minutes_file=None)
EVIDENCE_DEFAULTS = dict(outline_question_id='不适用', note=None, source_sha256=None, record_ids=[])


class ContractError(ValueError):
    pass


def split_ids(value):
    return [v.strip() for v in re.split(r'[\n；;、]+', str(value or '')) if v.strip()]


def unique(values):
    return list(dict.fromkeys(values))


def canonical(data):
    data = deepcopy(data)
    data.setdefault('interviews', [])
    if data.get('interview'):
        incoming = data['interview']
        data['interviews'] = [x for x in data['interviews'] if x['id'] != incoming['id']] + [incoming]
    data.setdefault('records', [])
    data.setdefault('evidence', [])
    for g in data.get('project', {}).get('groups', []): g['scope'] = g.get('scope') or ''
    for r in data['records']:
        if set(r) - RECORD_FIELDS: raise ContractError(f"{r.get('id')}: unknown record fields")
        for key in ['impact', 'management_request', 'historical_id']:
            r.setdefault(key, None)
            if r[key] == '': r[key] = None
        for key in ['collaborating_groups', 'departments', 'related_ids', 'interview_ids']:
            r.setdefault(key, [])
        r.setdefault('status', '待核实')
        r['verification'] = {**VERIFY_DEFAULTS, **r.get('verification', {})}
        for key in VERIFY_DEFAULTS:
            if key != 'priority' and r['verification'][key] == '': r['verification'][key] = None
    for i in data['interviews']:
        for key, value in INTERVIEW_DEFAULTS.items(): i.setdefault(key, deepcopy(value))
        for key in INTERVIEW_DEFAULTS:
            if i[key] == '': i[key] = None
    for e in data['evidence']:
        for key, value in EVIDENCE_DEFAULTS.items(): e.setdefault(key, deepcopy(value))
        for key in ['note','source_sha256']:
            if e[key] == '': e[key] = None
    return data


def validate(data):
    """Mechanical evidence-type gates, not a claim that assertions are true."""
    errors = []
    project = data.get('project', {})
    if not project.get('name') or not re.fullmatch(r'[A-Za-z0-9][A-Za-z0-9-]*', project.get('code', '')):
        errors.append('project.name and an ASCII project.code are required')
    groups = [g.get('name') for g in project.get('groups', [])]
    if not groups or any(not g or '待配置' in g for g in groups) or len(groups) != len(set(groups)):
        errors.append('project.groups must contain unique, configured group names')
    records = data.get('records', [])
    evidence = data.get('evidence', [])
    interviews = data.get('interviews', [])
    for label, rows in [('records', records), ('evidence', evidence), ('interviews', interviews)]:
        ids = [r.get('id') for r in rows]
        if any(not isinstance(i, str) or not i.strip() for i in ids) or len(ids) != len(set(ids)):
            errors.append(f'{label}: missing or duplicate IDs')
    ids = {r.get('id') for r in records}
    iids = {i.get('id') for i in interviews}
    emap = {e.get('id'): e for e in evidence}
    by_record = {i: [] for i in ids}
    for e in evidence:
        eid = e.get('id')
        for key in ['interview_id', 'speaker', 'source_file', 'locator', 'quote', 'purpose']:
            if not e.get(key): errors.append(f'{eid}: missing {key}')
        if e.get('interview_id') not in iids: errors.append(f'{eid}: unknown interview')
        if e.get('utterance_type') not in UTTERANCES: errors.append(f'{eid}: invalid utterance_type')
        if e.get('speaker_role') not in {'client', 'consultant', 'unknown'}:
            errors.append(f'{eid}: speaker_role must be client, consultant or unknown')
        for rid in e.get('record_ids', []):
            if rid not in ids: errors.append(f'{eid}: unknown record {rid}')
            else: by_record[rid].append(e)
    for r in records:
        rid = r.get('id')
        if not re.fullmatch(re.escape(project.get('code', '')) + r'-ISS-\d+', rid or ''):
            errors.append(f'{rid}: use a project-wide ID, e.g. PROJECT-ISS-0001')
        for key in ['domain', 'title', 'description']:
            if not r.get(key): errors.append(f'{rid}: missing {key}')
        if r.get('group') not in groups: errors.append(f'{rid}: unknown group')
        peers = r.get('collaborating_groups', [])
        if any(g not in groups or g == r.get('group') for g in peers) or len(peers) != len(set(peers)):
            errors.append(f'{rid}: invalid collaborating_groups')
        if r.get('nature') not in NATURES or r.get('status') not in STATUSES:
            errors.append(f'{rid}: invalid nature or status')
        if not r.get('interview_ids') or any(i not in iids for i in r.get('interview_ids', [])):
            errors.append(f'{rid}: missing or unknown source interview')
        if any(i not in ids or i == rid for i in r.get('related_ids', [])):
            errors.append(f'{rid}: invalid related_ids')
        if r.get('status') == '已合并' and not r.get('related_ids'):
            errors.append(f'{rid}: merged record needs a target ID')
        ver = r.get('verification', {})
        for key in ['focus', 'materials', 'completion_criteria']:
            if not ver.get(key): errors.append(f'{rid}: missing verification.{key}')
        if ver.get('priority') not in {'高', '中', '低', '待评估'}:
            errors.append(f'{rid}: invalid priority')
        for key in ['due_date', 'confirmed_date']:
            if ver.get(key):
                try: date.fromisoformat(ver[key])
                except (ValueError, TypeError): errors.append(f'{rid}: {key} must be YYYY-MM-DD')
        if r.get('status') in {'已确认', '已合并', '已排除'}:
            if not all(ver.get(k) for k in ['conclusion', 'confirmer', 'confirmed_date']):
                errors.append(f'{rid}: confirmation/disposition needs conclusion, confirmer and date')
        linked = by_record.get(rid, [])
        if not linked: errors.append(f'{rid}: no linked evidence')
        if any(e['interview_id'] not in r.get('interview_ids', []) for e in linked):
            errors.append(f'{rid}: evidence/source-interview mismatch')
        if r.get('nature') == '问题' and not any(e.get('utterance_type') == '现状陈述' for e in linked):
            errors.append(f'{rid}: questions, requests and suggestions alone cannot establish a problem')
        if r.get('management_request') and not any(e.get('speaker_role') == 'client' and
                e.get('utterance_type') in {'管理诉求', '明确决定', '明确安排'} for e in linked):
            errors.append(f'{rid}: management_request needs client request/decision evidence')
    for i in interviews:
        if i.get('date'):
            try: date.fromisoformat(i['date'])
            except (ValueError, TypeError): errors.append(f"{i.get('id')}: invalid interview date")
        hours = i.get('duration_hours')
        if hours is not None and (not isinstance(hours, (int, float)) or isinstance(hours, bool) or hours <= 0):
            errors.append(f"{i.get('id')}: duration_hours must be positive or null")
    minutes = data.get('minutes')
    if minutes:
        def check_refs(item, context):
            refs = item.get('evidence_ids', [])
            if not refs or any(e not in emap for e in refs):
                errors.append(f'{context}: missing/unknown evidence_ids'); return []
            return [emap[e] for e in refs]
        def section_check(section):
            if not section.get('title'): errors.append('minutes: section title is required')
            for item in section.get('items', []):
                linked = check_refs(item, 'minutes item')
                kind = item.get('kind')
                if not item.get('label') or not item.get('text') or kind not in KINDS:
                    errors.append('minutes item needs label, text and kind')
                if kind == 'fact' and not any(e.get('utterance_type') == '现状陈述' for e in linked):
                    errors.append('minutes fact needs a current-state statement')
                if kind == 'decision' and not any(e.get('utterance_type') == '明确决定' for e in linked):
                    errors.append('minutes decision needs explicit decision evidence')
                if kind == 'request' and not any(e.get('speaker_role') == 'client' and e.get('utterance_type') in {'管理诉求','明确决定','明确安排'} for e in linked):
                    errors.append('minutes request needs client request evidence')
                if kind == 'suggestion' and not any(e.get('utterance_type') == '顾问建议' for e in linked):
                    errors.append('minutes suggestion needs suggestion evidence')
            for sub in section.get('subsections', []): section_check(sub)
        for section in minutes.get('sections', []): section_check(section)
        for action in minutes.get('actions', []):
            linked = check_refs(action, 'action')
            if (action.get('owner') or action.get('when')) and not action.get('agreed'):
                errors.append('unagreed action cannot contain an assigned owner or deadline')
            if action.get('agreed') and not any(e.get('utterance_type') in {'明确决定', '明确安排'} for e in linked):
                errors.append('agreed action needs explicit decision/arrangement evidence')
    if errors: raise ContractError('\n'.join(errors))


def verify_sources(evidence, source_dir):
    root = Path(source_dir).resolve()
    cached, checked = {}, []
    for e in evidence:
        path = (root / e['source_file']).resolve()
        if not path.is_relative_to(root): raise ContractError('source_file escapes --sources-dir')
        if path not in cached:
            raw = path.read_bytes()
            cached[path] = (re.sub(r'\s+', '', raw.decode('utf-8-sig')), hashlib.sha256(raw).hexdigest())
        content, digest = cached[path]
        if re.sub(r'\s+', '', e['quote']) not in content:
            raise ContractError(f"{e['id']}: quote not found in the supplied source")
        if e.get('source_sha256') and e['source_sha256'] != digest:
            raise ContractError(f"{e['id']}: source hash changed")
        checked.append({'id': e['id'], 'source_file': e['source_file'], 'sha256': digest})
    return checked


def merge(existing, incoming):
    """Preserve business edits; update existing fields only through explicit patches."""
    incoming = canonical(incoming)
    if existing is None:
        if incoming.get('record_updates'): raise ContractError('record_updates needs an existing ledger')
        return incoming, []
    existing = canonical(existing)
    for key in ['name', 'code']:
        if existing['project'][key] != incoming['project'][key]: raise ContractError('project identity mismatch')
    state = deepcopy(existing)
    state['project'] = incoming['project']
    records = {r['id']: r for r in state['records']}
    changes = []
    for patch in incoming.get('record_updates', []):
        rid, updates = patch.get('id'), patch.get('changes', {})
        if rid not in records or not patch.get('reason') or not updates:
            raise ContractError('record_updates needs an existing id, reason and changes')
        if set(updates) - (RECORD_FIELDS - {'id', 'historical_id'}): raise ContractError('unknown/immutable update field')
        before = deepcopy(records[rid])
        for key, value in updates.items():
            if key == 'verification':
                if set(value) - set(VERIFY_DEFAULTS): raise ContractError('unknown verification update field')
                records[rid][key].update(value)
            else: records[rid][key] = value
        changes.append({'id': rid, 'reason': patch['reason'], 'before': before, 'after': deepcopy(records[rid])})
    for r in incoming['records']:
        if r['id'] in records:
            old = records[r['id']]
            for key in RECORD_FIELDS - {'interview_ids', 'related_ids'}:
                if r.get(key) != old.get(key): raise ContractError(f"{r['id']}: use record_updates to change {key}")
            for key in ['interview_ids', 'related_ids']: old[key] = unique(old.get(key, []) + r.get(key, []))
        else: records[r['id']] = deepcopy(r)
    state['records'] = list(records.values())
    for field in ['evidence', 'interviews']:
        indexed = {x['id']: x for x in state[field]}
        for item in incoming[field]:
            if item['id'] in indexed and indexed[item['id']] != item:
                if field == 'interviews':
                    old = indexed[item['id']]
                    for key,value in item.items():
                        if value in (None, [], ''): continue
                        if key == 'minutes_file': old[key] = value
                        elif old.get(key) in (None, [], ''): old[key] = value
                        elif old.get(key) != value: raise ContractError(f"{item['id']}: conflicting interview metadata: {key}")
                    continue
                raise ContractError(f"{item['id']}: conflicting immutable evidence entry")
            indexed[item['id']] = deepcopy(item)
        state[field] = list(indexed.values())
    state['minutes'] = incoming.get('minutes')
    state['interview'] = incoming.get('interview')
    for e in incoming['evidence']:
        for rid in e.get('record_ids', []):
            if rid in records:
                records[rid]['interview_ids'] = unique(records[rid]['interview_ids'] + [e['interview_id']])
    return state, changes


def counts(records):
    active = [r for r in records if r['status'] not in {'已合并', '已排除'}]
    return [sum(r['nature'] == '问题' for r in active),
            sum(r['nature'] == '待核实线索' for r in active),
            sum(r['nature'] in {'综合议题', '管理诉求', '实施风险'} for r in active),
            sum(r['nature'] == '问题' and r['status'] == '已确认' for r in records)]
