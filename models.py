"""Database models for SLT Project Management."""
import json
from datetime import datetime, date
from flask_sqlalchemy import SQLAlchemy

db = SQLAlchemy()


def _d(v):
    """Serialise a date/datetime to ISO text (or None)."""
    if v is None:
        return None
    return v.isoformat()


class Project(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    code = db.Column(db.String(50), default='')
    name = db.Column(db.String(255), nullable=False)
    manager = db.Column(db.String(255), default='')
    description = db.Column(db.Text, default='')
    currency = db.Column(db.String(10), default='LKR')
    status_date = db.Column(db.Date, default=date.today)
    source_name = db.Column(db.String(255), default='')
    pod_template = db.Column(db.LargeBinary)          # original .pod bytes (for .pod export)
    created_at = db.Column(db.DateTime, default=datetime.now)
    updated_at = db.Column(db.DateTime, default=datetime.now, onupdate=datetime.now)

    tasks = db.relationship('Task', backref='project', cascade='all, delete-orphan', order_by='Task.seq')
    snapshots = db.relationship('Snapshot', backref='project', cascade='all, delete-orphan', order_by='Snapshot.status_date')
    scopes = db.relationship('Scope', backref='project', cascade='all, delete-orphan', order_by='Scope.seq')
    planned = db.relationship('PlannedPoint', backref='project', cascade='all, delete-orphan', order_by='PlannedPoint.month')

    def to_dict(self):
        return dict(id=self.id, code=self.code, name=self.name, manager=self.manager, description=self.description,
                    currency=self.currency, status_date=_d(self.status_date), source_name=self.source_name,
                    has_pod=self.pod_template is not None, created_at=_d(self.created_at), updated_at=_d(self.updated_at))


class Task(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    project_id = db.Column(db.Integer, db.ForeignKey('project.id'), nullable=False, index=True)
    uid = db.Column(db.Integer, nullable=False)          # ProjectLibre UID (stable key)
    seq = db.Column(db.Integer, nullable=False)          # row order (ProjectLibre ID)
    name = db.Column(db.String(500), default='')
    level = db.Column(db.Integer, default=1)             # outline level
    wbs = db.Column(db.String(50), default='')
    is_summary = db.Column(db.Boolean, default=False)
    is_milestone = db.Column(db.Boolean, default=False)
    start = db.Column(db.DateTime)
    finish = db.Column(db.DateTime)
    bl_start = db.Column(db.DateTime)
    bl_finish = db.Column(db.DateTime)
    duration_h = db.Column(db.Float, default=8)
    pct = db.Column(db.Float, default=0)                 # 0..100
    text1 = db.Column(db.String(50), default='')         # "1" = not counted in progress %
    text2 = db.Column(db.Text, default='')               # remarks
    notes = db.Column(db.Text, default='')
    predecessors = db.Column(db.Text, default='[]')      # JSON [{uid,type,lag}]
    constraint_type = db.Column(db.Integer, default=0)
    constraint_date = db.Column(db.DateTime)

    def to_dict(self):
        return dict(id=self.id, uid=self.uid, seq=self.seq, name=self.name, level=self.level, wbs=self.wbs,
                    is_summary=self.is_summary, is_milestone=self.is_milestone, start=_d(self.start),
                    finish=_d(self.finish), bl_start=_d(self.bl_start), bl_finish=_d(self.bl_finish),
                    duration_h=self.duration_h, pct=self.pct, text1=self.text1 or '', text2=self.text2 or '',
                    notes=self.notes or '', predecessors=json.loads(self.predecessors or '[]'),
                    constraint_type=self.constraint_type, constraint_date=_d(self.constraint_date))


class Snapshot(db.Model):
    """A recorded status update (e.g. 'October - 2026'): a frozen copy of every task."""
    id = db.Column(db.Integer, primary_key=True)
    project_id = db.Column(db.Integer, db.ForeignKey('project.id'), nullable=False, index=True)
    label = db.Column(db.String(100))
    status_date = db.Column(db.Date, nullable=False)
    source = db.Column(db.String(255), default='')
    created_at = db.Column(db.DateTime, default=datetime.now)
    data = db.Column(db.Text, nullable=False)            # JSON list of task dicts

    def to_dict(self, with_data=False):
        d = dict(id=self.id, label=self.label, status_date=_d(self.status_date), source=self.source,
                 created_at=_d(self.created_at))
        if with_data:
            d['tasks'] = json.loads(self.data)
        return d


class PlannedPoint(db.Model):
    """Manually planned overall progress for a month (the 'Planned Progress' input)."""
    id = db.Column(db.Integer, primary_key=True)
    project_id = db.Column(db.Integer, db.ForeignKey('project.id'), nullable=False, index=True)
    month = db.Column(db.String(7), nullable=False)      # YYYY-MM
    pct = db.Column(db.Float, nullable=False)            # 0..100

    def to_dict(self):
        return dict(id=self.id, month=self.month, pct=self.pct)


class Scope(db.Model):
    """A CAPEX budget scope / tender with payment terms tied to milestone tasks."""
    id = db.Column(db.Integer, primary_key=True)
    project_id = db.Column(db.Integer, db.ForeignKey('project.id'), nullable=False, index=True)
    seq = db.Column(db.Integer, default=0)
    name = db.Column(db.String(500), default='')
    tender_no = db.Column(db.String(100), default='')
    component_uid = db.Column(db.Integer)                # component (summary task) the scope belongs to
    reserved = db.Column(db.Float, default=0)
    alloc_goods = db.Column(db.Float, default=0)
    alloc_services = db.Column(db.Float, default=0)
    g_grn = db.Column(db.Float, default=0)               # fractions 0..1
    g_pac = db.Column(db.Float, default=0)
    g_fac = db.Column(db.Float, default=0)
    s_grn = db.Column(db.Float, default=0)
    s_pac = db.Column(db.Float, default=0)
    s_fac = db.Column(db.Float, default=0)
    grn_uid = db.Column(db.Integer)
    pac_uid = db.Column(db.Integer)
    fac_uid = db.Column(db.Integer)
    paid_actual = db.Column(db.Float)                    # optional: amount finance has actually paid
    terms_text = db.Column(db.Text, default='')

    FIELDS = ['seq', 'name', 'tender_no', 'component_uid', 'reserved', 'alloc_goods', 'alloc_services', 'g_grn',
              'g_pac', 'g_fac', 's_grn', 's_pac', 's_fac', 'grn_uid', 'pac_uid', 'fac_uid', 'paid_actual', 'terms_text']

    def to_dict(self):
        d = {f: getattr(self, f) for f in self.FIELDS}
        d['id'] = self.id
        return d
