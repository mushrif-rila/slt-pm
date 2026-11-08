"""SLT Project Management — web application.

Run:   python app.py                 (opens on http://127.0.0.1:5000)
CLI:   python app.py import FILE...  (create a project from files: .pod / .xml / .xlsx, oldest first)
"""
import calendar
import io
import json
import os
import re
import sys
from datetime import date, datetime

from flask import Flask, abort, jsonify, request, send_file, send_from_directory

from models import PlannedPoint, Project, Scope, Snapshot, Task, db
from services import calc
from services.excel_io import export_workbook, read_workbook
from services.mspdi import FormatError, build_pod, build_xml, parse_dt, read_project_file
from services.pdf_report import build_pdf

BASE = os.path.dirname(os.path.abspath(__file__))
app = Flask(__name__, static_folder='static', static_url_path='/static')
app.config['SQLALCHEMY_DATABASE_URI'] = os.environ.get('SLT_PM_DB', 'sqlite:///' + os.path.join(BASE, 'slt_pm.db'))
app.config['MAX_CONTENT_LENGTH'] = 100 * 1024 * 1024
db.init_app(app)
with app.app_context():
    db.create_all()


# ===================================================================== helpers
def err(msg, code=400):
    return jsonify(error=msg), code


@app.errorhandler(FormatError)
def _fmt_err(e):
    return err(str(e))


def get_project(pid):
    p = db.session.get(Project, pid)
    if not p:
        abort(404)
    return p


def live_tasks(p):
    return [t.to_dict() for t in p.tasks]


def _dt(v):
    if v in (None, ''):
        return None
    if isinstance(v, datetime):
        return v
    return parse_dt(str(v))


def save_tasks(p, tasks):
    """Replace the project's live plan with the given task dicts."""
    Task.query.filter_by(project_id=p.id).delete()
    for i, t in enumerate(tasks, start=1):
        db.session.add(Task(
            project_id=p.id, uid=int(t['uid']), seq=i, name=t.get('name') or '', level=int(t.get('level') or 1),
            wbs=t.get('wbs') or '', is_summary=bool(t.get('is_summary')), is_milestone=bool(t.get('is_milestone')),
            start=_dt(t.get('start')), finish=_dt(t.get('finish')), bl_start=_dt(t.get('bl_start')),
            bl_finish=_dt(t.get('bl_finish')), duration_h=float(t.get('duration_h') or 0), pct=float(t.get('pct') or 0),
            text1=str(t.get('text1') or ''), text2=t.get('text2') or '', notes=t.get('notes') or '',
            predecessors=json.dumps(t.get('predecessors') or []), constraint_type=int(t.get('constraint_type') or 0),
            constraint_date=_dt(t.get('constraint_date'))))
    p.updated_at = datetime.now()


def snap_dict(s):
    d = s.to_dict(with_data=True)
    return d


def scopes_of(p):
    return [s.to_dict() for s in p.scopes]


def planned_of(p):
    return [x.to_dict() for x in p.planned]


def resolve_view(p, view):
    """Returns (tasks, status_date, label, previous_tasks) for 'live' or a snapshot id."""
    snaps = list(p.snapshots)
    if view in (None, '', 'latest'):
        view = str(snaps[-1].id) if snaps else 'live'
    if view == 'live':
        prev = json.loads(snaps[-1].data) if snaps else None
        return live_tasks(p), p.status_date or date.today(), 'Current Plan', prev
    s = db.session.get(Snapshot, int(view))
    if not s or s.project_id != p.id:
        abort(404)
    idx = snaps.index(s)
    prev = json.loads(snaps[idx - 1].data) if idx > 0 else None
    return json.loads(s.data), s.status_date, s.label, prev


def analysis_for(p, view):
    tasks, sd, label, prev = resolve_view(p, view)
    a = calc.analyze(tasks, sd, scopes_of(p), prev, calc.planned_for(planned_of(p), sd))
    a['label'] = label
    return a


def trend_of(p):
    snaps = [dict(id=s.id, label=s.label, status_date=s.status_date.isoformat(), tasks=json.loads(s.data))
             for s in p.snapshots]
    return calc.trend(snaps, scopes_of(p), planned_of(p))


def month_end(y, m):
    return date(y, m, calendar.monthrange(y, m)[1])


def guess_status_date(filename, info=None):
    m = re.search(r'(20\d\d)[-_ ](0[1-9]|1[0-2])(?!\d)', filename or '')
    if m:
        return month_end(int(m.group(1)), int(m.group(2)))
    if info:
        for k in ('status_date', 'current_date'):
            if info.get(k):
                return info[k].date()
    return date.today()


def upsert_snapshot(p, tasks, sd, label=None, source=''):
    label = label or calc.month_label(sd)
    s = Snapshot.query.filter_by(project_id=p.id, label=label).first()
    if s:
        s.status_date, s.data, s.source, s.created_at = sd, json.dumps(tasks), source, datetime.now()
    else:
        s = Snapshot(project_id=p.id, label=label, status_date=sd, data=json.dumps(tasks), source=source)
        db.session.add(s)
    return s


def merge_progress(existing, incoming):
    """'Progress only' update: keep the live structure, take % / dates / remarks of matching UIDs."""
    by_uid = {t['uid']: t for t in incoming}
    out = []
    for t in existing:
        n = by_uid.get(t['uid'])
        if n:
            t = dict(t)
            for k in ('pct', 'start', 'finish', 'text2', 'text1', 'duration_h'):
                if n.get(k) not in (None, ''):
                    t[k] = n[k]
        out.append(t)
    return calc.rollup(out)


def keep_baselines(existing, incoming):
    old = {t['uid']: t for t in existing}
    for t in incoming:
        o = old.get(t['uid'])
        if o and not t.get('bl_start') and o.get('bl_start'):
            t['bl_start'], t['bl_finish'] = o['bl_start'], o['bl_finish']
    return incoming


def parse_upload(fs):
    raw = fs.read()
    name = fs.filename or 'upload'
    ext = os.path.splitext(name)[1].lower()
    if ext in ('.xlsx', '.xlsm'):
        return dict(kind='xlsx', name=name, data=read_workbook(raw))
    if ext in ('.pod', '.xml', '.mspdi') or raw[:2] == b'\xac\xed' or b'<Project' in raw[:2000]:
        info, tasks = read_project_file(raw)
        return dict(kind='pod' if raw[:2] == b'\xac\xed' else 'xml', name=name, info=info, tasks=tasks,
                    raw=raw if raw[:2] == b'\xac\xed' else None)
    raise FormatError(f'{name}: unsupported file type. Upload a .pod, .xml or .xlsx file.')


def apply_files(p, files, mode='replace', status_date=None, label=None):
    """Apply uploaded files to a project (oldest status first). Returns list of created snapshot labels."""
    parsed = [parse_upload(f) for f in files]
    items = []
    for u in parsed:
        if u['kind'] == 'xlsx':
            for sh in u['data']['sheets']:
                items.append(dict(sd=sh['status_date'] or guess_status_date(u['name']), tasks=sh['tasks'],
                                  src=f"{u['name']} › {sh['sheet']}", raw=None, label=None))
            if u['data']['scopes'] and not p.scopes:
                for i, s in enumerate(u['data']['scopes']):
                    db.session.add(Scope(project_id=p.id, seq=i, **s))
            for pl in u['data']['planned']:
                x = PlannedPoint.query.filter_by(project_id=p.id, month=pl['month']).first()
                if x:
                    x.pct = pl['pct']
                else:
                    db.session.add(PlannedPoint(project_id=p.id, month=pl['month'], pct=round(pl['pct'], 2)))
        else:
            items.append(dict(sd=guess_status_date(u['name'], u['info']), tasks=u['tasks'], src=u['name'],
                              raw=u['raw'], label=None, info=u['info']))
    if not items:
        raise FormatError('Nothing to import.')
    if len(items) == 1 and status_date:
        items[0]['sd'] = status_date
        items[0]['label'] = label or None
    items.sort(key=lambda x: x['sd'])
    labels = []
    for it in items:
        cur = live_tasks(p)
        if mode == 'progress' and cur:
            new = merge_progress(cur, it['tasks'])
        else:
            new = keep_baselines(cur, [dict(t) for t in it['tasks']])
        save_tasks(p, new)
        db.session.flush()
        if it.get('raw'):
            p.pod_template = it['raw']
        if it.get('info') and not p.manager:
            p.manager = it['info'].get('manager') or ''
        p.status_date = it['sd']
        s = upsert_snapshot(p, new, it['sd'], it.get('label'), it['src'])
        if s.label not in labels:
            labels.append(s.label)
    p.source_name = ', '.join(sorted({u['name'] for u in parsed}))[:250]
    return labels


def new_project_from_files(files, fields):
    parsed_name = None
    files = list(files)
    head = files[0]
    pos = head.stream.tell()
    first = parse_upload(head)
    head.stream.seek(pos)
    if first['kind'] == 'xlsx':
        parsed_name = first['data']['name']
    else:
        parsed_name = first['info']['name']
        fields.setdefault('manager', first['info'].get('manager') or '')
    code = fields.get('code') or ''
    if not code:
        m = re.match(r'(\d{6})[_-]', first['name'])
        code = m.group(1) if m else ''
    p = Project(name=fields.get('name') or parsed_name or 'Imported project', code=code,
                manager=fields.get('manager') or '', description=fields.get('description') or '',
                currency=fields.get('currency') or 'LKR')
    db.session.add(p); db.session.flush()
    for f in files:
        f.stream.seek(0)
    labels = apply_files(p, files, 'replace')
    return p, labels


def project_card(p):
    d = p.to_dict()
    snaps = list(p.snapshots)
    if p.tasks or snaps:
        a = analysis_for(p, 'latest')
        d.update(kpi=a['kpi'], health=a['health'], budget=a['budget']['totals'], view_label=a['label'],
                 components=[dict(name=c['name'], actual=c['actual'], status=c['status']) for c in a['components']])
        d['spark'] = [dict(label=t['label'], actual=t['actual']) for t in trend_of(p)][-8:]
    d['task_count'] = len(p.tasks)
    d['snapshot_count'] = len(snaps)
    d['scope_count'] = len(p.scopes)
    return d


def safe_name(*parts):
    s = '_'.join(str(x) for x in parts if x)
    return re.sub(r'[^A-Za-z0-9._-]+', '_', s).strip('_')[:120]


# ===================================================================== pages
@app.route('/')
def index():
    return send_from_directory(os.path.join(BASE, 'static'), 'index.html')


# ===================================================================== projects
@app.get('/api/projects')
def list_projects():
    return jsonify([project_card(p) for p in Project.query.order_by(Project.updated_at.desc()).all()])


@app.post('/api/projects')
def create_project():
    d = request.get_json(force=True)
    if not (d.get('name') or '').strip():
        return err('Project name is required.')
    p = Project(name=d['name'].strip(), code=d.get('code', ''), manager=d.get('manager', ''),
                description=d.get('description', ''), currency=d.get('currency') or 'LKR',
                status_date=calc.to_date(d.get('status_date')) or date.today())
    db.session.add(p); db.session.flush()
    start = calc.to_date(d.get('start_date')) or date.today()
    s = datetime(start.year, start.month, start.day, 8)
    f = datetime.combine(calc.add_workdays(start, 5), datetime.min.time()).replace(hour=17)
    tasks = [dict(uid=1, name=p.name, level=1, start=s.isoformat(), finish=f.isoformat(), duration_h=40, pct=0, text1='1'),
             dict(uid=2, name='Project Initiation', level=2, start=s.isoformat(), finish=f.isoformat(), duration_h=40, pct=0, text1='1'),
             dict(uid=3, name='G1 Approval', level=3, start=s.isoformat(), finish=f.isoformat(), duration_h=40, pct=0)]
    save_tasks(p, calc.rollup(tasks))
    db.session.commit()
    return jsonify(p.to_dict()), 201


@app.post('/api/projects/import')
def import_project():
    files = request.files.getlist('files')
    if not files:
        return err('Choose at least one .pod, .xml or .xlsx file.')
    p, labels = new_project_from_files(files, {k: v for k, v in request.form.items() if v})
    db.session.commit()
    return jsonify(project=p.to_dict(), snapshots=labels), 201


@app.get('/api/projects/<int:pid>')
def get_project_api(pid):
    p = get_project(pid)
    d = p.to_dict()
    d['snapshots'] = [s.to_dict() for s in p.snapshots]
    d['planned'] = planned_of(p)
    return jsonify(d)


@app.put('/api/projects/<int:pid>')
def update_project(pid):
    p = get_project(pid)
    d = request.get_json(force=True)
    for k in ('name', 'code', 'manager', 'description', 'currency'):
        if k in d:
            setattr(p, k, d[k] or '')
    if d.get('status_date'):
        p.status_date = calc.to_date(d['status_date'])
    db.session.commit()
    return jsonify(p.to_dict())


@app.delete('/api/projects/<int:pid>')
def delete_project(pid):
    p = get_project(pid)
    db.session.delete(p); db.session.commit()
    return jsonify(ok=True)


@app.get('/api/projects/<int:pid>/dashboard')
def dashboard(pid):
    p = get_project(pid)
    a = analysis_for(p, request.args.get('view'))
    return jsonify(project=p.to_dict(), analysis=a, trend=trend_of(p),
                   snapshots=[s.to_dict() for s in p.snapshots], scopes=scopes_of(p), planned=planned_of(p))


# ===================================================================== status updates
@app.post('/api/projects/<int:pid>/updates')
def upload_update(pid):
    p = get_project(pid)
    files = request.files.getlist('files')
    if not files:
        return err('Choose at least one .pod, .xml or .xlsx file.')
    sd = calc.to_date(request.form.get('status_date')) if request.form.get('status_date') else None
    labels = apply_files(p, files, request.form.get('mode', 'replace'), sd, request.form.get('label') or None)
    db.session.commit()
    return jsonify(ok=True, snapshots=labels)


@app.post('/api/projects/<int:pid>/snapshots')
def create_snapshot(pid):
    p = get_project(pid)
    d = request.get_json(force=True)
    sd = calc.to_date(d.get('status_date')) or p.status_date or date.today()
    s = upsert_snapshot(p, live_tasks(p), sd, d.get('label') or None, 'Recorded in SLT Project Management')
    p.status_date = sd
    db.session.commit()
    return jsonify(s.to_dict()), 201


@app.put('/api/snapshots/<int:sid>')
def edit_snapshot(sid):
    s = db.session.get(Snapshot, sid) or abort(404)
    d = request.get_json(force=True)
    if d.get('label'):
        s.label = d['label']
    if d.get('status_date'):
        s.status_date = calc.to_date(d['status_date'])
    db.session.commit()
    return jsonify(s.to_dict())


@app.delete('/api/snapshots/<int:sid>')
def delete_snapshot(sid):
    s = db.session.get(Snapshot, sid) or abort(404)
    db.session.delete(s); db.session.commit()
    return jsonify(ok=True)


@app.post('/api/snapshots/<int:sid>/restore')
def restore_snapshot(sid):
    s = db.session.get(Snapshot, sid) or abort(404)
    p = s.project
    save_tasks(p, json.loads(s.data))
    p.status_date = s.status_date
    db.session.commit()
    return jsonify(ok=True)


# ===================================================================== tasks (live plan)
@app.get('/api/projects/<int:pid>/tasks')
def get_tasks(pid):
    p = get_project(pid)
    a = calc.analyze(live_tasks(p), p.status_date or date.today(), scopes_of(p))
    return jsonify(tasks=a['tasks'], status_date=(p.status_date or date.today()).isoformat())


def _commit_rollup(p, tasks):
    save_tasks(p, calc.rollup(tasks))
    db.session.commit()
    return get_tasks(p.id)


def _block_end(tasks, i):
    j = i + 1
    while j < len(tasks) and tasks[j]['level'] > tasks[i]['level']:
        j += 1
    return j


@app.put('/api/tasks/<int:tid>')
def update_task(tid):
    t = db.session.get(Task, tid) or abort(404)
    p = t.project
    d = request.get_json(force=True)
    tasks = live_tasks(p)
    me = next(x for x in tasks if x['id'] == tid)
    for k in ('name', 'text1', 'text2', 'notes'):
        if k in d:
            me[k] = (d[k] or '').strip() if isinstance(d[k], str) else d[k]
    if 'pct' in d:
        me['pct'] = max(0.0, min(100.0, float(d['pct'] or 0)))
    if 'is_milestone' in d:
        me['is_milestone'] = bool(d['is_milestone'])
    for k in ('start', 'finish', 'bl_start', 'bl_finish'):
        if k in d:
            v = _dt(d[k])
            if v and k in ('start', 'bl_start') and v.hour == 0:
                v = v.replace(hour=8)
            if v and k in ('finish', 'bl_finish') and v.hour == 0:
                v = v.replace(hour=17)
            me[k] = v.isoformat() if v else None
    if 'duration_days' in d and me.get('start'):
        s = calc.to_date(me['start'])
        f = calc.add_workdays(s, float(d['duration_days'] or 1))
        me['finish'] = datetime(f.year, f.month, f.day, 17).isoformat()
    if 'predecessors' in d:
        from services.excel_io import _parse_preds
        me['predecessors'] = _parse_preds(d['predecessors']) if isinstance(d['predecessors'], str) else d['predecessors']
    if me.get('start') and me.get('finish'):
        s, f = calc.to_date(me['start']), calc.to_date(me['finish'])
        if f < s:
            return err('Finish date is before start date.')
        me['duration_h'] = max(calc.workdays(s, f), 0 if me.get('is_milestone') else 1) * 8
    return _commit_rollup(p, tasks)


@app.post('/api/projects/<int:pid>/tasks')
def add_task(pid):
    p = get_project(pid)
    d = request.get_json(force=True)
    tasks = live_tasks(p)
    uid = max([t['uid'] for t in tasks] + [0]) + 1
    sd = p.status_date or date.today()
    new = dict(uid=uid, name=d.get('name') or 'New task', level=1, pct=0, text1='', text2='',
               start=datetime(sd.year, sd.month, sd.day, 8).isoformat(),
               finish=datetime.combine(calc.add_workdays(sd, 5), datetime.min.time()).replace(hour=17).isoformat(),
               duration_h=40, predecessors=[])
    after = d.get('after_id')
    if after:
        i = next((k for k, t in enumerate(tasks) if t['id'] == after), None)
        if i is None:
            return err('Row not found.')
        ref = tasks[i]
        if d.get('as_child'):
            new['level'] = ref['level'] + 1
            pos = i + 1
        else:
            new['level'] = ref['level']
            pos = _block_end(tasks, i)
        if ref.get('start') and not ref.get('is_summary'):
            new['start'], new['finish'], new['duration_h'] = ref['start'], ref['finish'], ref.get('duration_h') or 40
        tasks.insert(pos, new)
    else:
        new['level'] = 2 if tasks and tasks[0]['level'] == 1 and len([t for t in tasks if t['level'] == 1]) == 1 else 1
        tasks.append(new)
    return _commit_rollup(p, tasks)


@app.delete('/api/tasks/<int:tid>')
def delete_task(tid):
    t = db.session.get(Task, tid) or abort(404)
    p = t.project
    tasks = live_tasks(p)
    i = next(k for k, x in enumerate(tasks) if x['id'] == tid)
    j = _block_end(tasks, i)
    gone = {x['uid'] for x in tasks[i:j]}
    tasks = tasks[:i] + tasks[j:]
    for x in tasks:
        x['predecessors'] = [q for q in x.get('predecessors') or [] if q['uid'] not in gone]
    return _commit_rollup(p, tasks)


@app.post('/api/tasks/<int:tid>/move')
def move_task(tid):
    t = db.session.get(Task, tid) or abort(404)
    p = t.project
    action = request.get_json(force=True).get('action')
    tasks = live_tasks(p)
    i = next(k for k, x in enumerate(tasks) if x['id'] == tid)
    j = _block_end(tasks, i)
    block = tasks[i:j]
    if action == 'indent':
        if i == 0 or tasks[i - 1]['level'] < block[0]['level']:
            return err('Cannot indent: no task above at the same level.')
        for x in block:
            x['level'] += 1
    elif action == 'outdent':
        if block[0]['level'] <= 1:
            return err('Already at the top level.')
        for x in block:          # following siblings become its children (same as ProjectLibre / MS Project)
            x['level'] -= 1
    elif action in ('up', 'down'):
        lvl = block[0]['level']
        if action == 'up':
            k = i - 1
            while k >= 0 and tasks[k]['level'] > lvl:
                k -= 1
            if k < 0 or tasks[k]['level'] < lvl:
                return err('Already first in its group.')
            tasks = tasks[:k] + block + tasks[k:i] + tasks[j:]
        else:
            if j >= len(tasks) or tasks[j]['level'] < lvl:
                return err('Already last in its group.')
            e = _block_end(tasks, j)
            tasks = tasks[:i] + tasks[j:e] + block + tasks[e:]
    else:
        return err('Unknown action.')
    return _commit_rollup(p, tasks)


@app.post('/api/projects/<int:pid>/baseline')
def set_baseline(pid):
    p = get_project(pid)
    only_empty = request.get_json(force=True).get('only_empty', False)
    tasks = live_tasks(p)
    for t in tasks:
        if only_empty and t.get('bl_start'):
            continue
        t['bl_start'], t['bl_finish'] = t.get('start'), t.get('finish')
    save_tasks(p, tasks)
    db.session.commit()
    return jsonify(ok=True)


# ===================================================================== planned progress
@app.put('/api/projects/<int:pid>/planned')
def put_planned(pid):
    p = get_project(pid)
    rows = request.get_json(force=True)
    PlannedPoint.query.filter_by(project_id=p.id).delete()
    for r in rows:
        if r.get('month') and r.get('pct') not in (None, ''):
            db.session.add(PlannedPoint(project_id=p.id, month=r['month'][:7], pct=float(r['pct'])))
    db.session.commit()
    return jsonify(planned_of(p))


# ===================================================================== budget scopes
def _scope_vals(d):
    out = {}
    for f in Scope.FIELDS:
        if f not in d:
            continue
        v = d[f]
        if f in ('name', 'tender_no', 'terms_text'):
            out[f] = v or ''
        elif f in ('component_uid', 'grn_uid', 'pac_uid', 'fac_uid', 'seq'):
            out[f] = int(v) if v not in (None, '') else None
        else:
            out[f] = float(v) if v not in (None, '') else (None if f == 'paid_actual' else 0)
    return out


@app.post('/api/projects/<int:pid>/scopes')
def add_scope(pid):
    p = get_project(pid)
    vals = _scope_vals(request.get_json(force=True))
    vals.setdefault('seq', len(p.scopes))
    s = Scope(project_id=p.id, **vals)
    db.session.add(s); db.session.commit()
    return jsonify(s.to_dict()), 201


@app.put('/api/scopes/<int:sid>')
def edit_scope(sid):
    s = db.session.get(Scope, sid) or abort(404)
    for k, v in _scope_vals(request.get_json(force=True)).items():
        setattr(s, k, v)
    db.session.commit()
    return jsonify(s.to_dict())


@app.delete('/api/scopes/<int:sid>')
def delete_scope(sid):
    s = db.session.get(Scope, sid) or abort(404)
    db.session.delete(s); db.session.commit()
    return jsonify(ok=True)


# ===================================================================== exports
def _view_tasks(p):
    tasks, sd, label, _ = resolve_view(p, request.args.get('view'))
    return tasks, sd, label


@app.get('/api/projects/<int:pid>/export/<fmt>')
def export(pid, fmt):
    p = get_project(pid)
    tasks, sd, label = _view_tasks(p)
    pd = p.to_dict(); pd['status_date'] = sd.isoformat()
    base = safe_name(p.code, p.name, label)
    if fmt == 'xml':
        return send_file(io.BytesIO(build_xml(pd, tasks)), mimetype='application/xml', as_attachment=True,
                         download_name=f'{base}.xml')
    if fmt == 'pod':
        if not p.pod_template:
            return err('No ProjectLibre .pod has been uploaded for this project yet, so a .pod cannot be built. '
                       'Download the XML instead — ProjectLibre opens it directly (File > Open) and can save it as .pod.')
        data = build_pod(p.pod_template, build_xml(pd, tasks) + b'\n')
        return send_file(io.BytesIO(data), mimetype='application/octet-stream', as_attachment=True,
                         download_name=f'{base}.pod')
    if fmt == 'xlsx':
        a = analysis_for(p, request.args.get('view'))
        snaps = [dict(id=s.id, label=s.label, status_date=s.status_date.isoformat(), tasks=json.loads(s.data))
                 for s in p.snapshots]
        live_pd = p.to_dict()
        data = export_workbook(live_pd, live_tasks(p), snaps, a, trend_of(p), scopes_of(p), planned_of(p))
        return send_file(io.BytesIO(data), as_attachment=True, download_name=f'{base}_Tracker.xlsx',
                         mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
    if fmt == 'pdf':
        a = analysis_for(p, request.args.get('view'))
        tr = [t for t in trend_of(p) if t['status_date'] <= a['status_date']][-10:]
        data = build_pdf(pd, a, tr)
        return send_file(io.BytesIO(data), mimetype='application/pdf', as_attachment=True,
                         download_name=f'{base}_Status.pdf')
    return err('Unknown format.')


# ===================================================================== CLI
class _LocalFile:
    def __init__(self, path):
        self.filename = os.path.basename(path)
        self.stream = open(path, 'rb')

    def read(self):
        return self.stream.read()


def cli_import(paths):
    with app.app_context():
        files = [_LocalFile(x) for x in paths]
        p, labels = new_project_from_files(files, {})
        db.session.commit()
        print(f'Created project #{p.id} "{p.name}" with status updates: {", ".join(labels)}')


if __name__ == '__main__':
    if len(sys.argv) > 2 and sys.argv[1] == 'import':
        cli_import(sys.argv[2:])
    else:
        host = os.environ.get('SLT_PM_HOST', '127.0.0.1')
        port = int(os.environ.get('SLT_PM_PORT', '5000'))
        print(f'SLT Project Management running on http://{host}:{port}  (Ctrl+C to stop)')
        try:
            from waitress import serve          # production-grade server, works on Windows / macOS / Linux
            serve(app, host=host, port=port, threads=8, max_request_body_size=100 * 1024 * 1024)
        except ImportError:
            app.run(host=host, port=port, debug=False)
