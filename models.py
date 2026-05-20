from datetime import datetime
from flask_sqlalchemy import SQLAlchemy
from flask_login import UserMixin
from werkzeug.security import generate_password_hash, check_password_hash

db = SQLAlchemy()


class User(UserMixin, db.Model):
    __tablename__ = "users"
    id = db.Column(db.Integer, primary_key=True)
    username = db.Column(db.String(64), unique=True, nullable=False, index=True)
    full_name = db.Column(db.String(128), nullable=False)
    password_hash = db.Column(db.String(255), nullable=False)
    default_billing_rate = db.Column(db.Float, nullable=False, default=0.0)
    is_admin = db.Column(db.Boolean, default=False, nullable=False)
    created_at = db.Column(db.DateTime, default=datetime.utcnow)

    memberships = db.relationship(
        "ProjectMember", backref="user", cascade="all, delete-orphan"
    )
    time_entries = db.relationship(
        "TimeEntry", backref="user", cascade="all, delete-orphan"
    )

    def set_password(self, password: str) -> None:
        self.password_hash = generate_password_hash(password)

    def check_password(self, password: str) -> bool:
        return check_password_hash(self.password_hash, password)


class Project(db.Model):
    __tablename__ = "projects"
    id = db.Column(db.Integer, primary_key=True)
    name = db.Column(db.String(128), nullable=False)
    client_name = db.Column(db.String(128), nullable=False)
    client_address = db.Column(db.Text, default="")
    default_billing_rate = db.Column(db.Float, nullable=False, default=0.0)
    archived = db.Column(db.Boolean, default=False, nullable=False)
    created_at = db.Column(db.DateTime, default=datetime.utcnow)

    members = db.relationship(
        "ProjectMember", backref="project", cascade="all, delete-orphan"
    )
    time_entries = db.relationship(
        "TimeEntry", backref="project", cascade="all, delete-orphan"
    )
    invoices = db.relationship(
        "Invoice", backref="project", cascade="all, delete-orphan"
    )


class ProjectMember(db.Model):
    __tablename__ = "project_members"
    id = db.Column(db.Integer, primary_key=True)
    project_id = db.Column(
        db.Integer, db.ForeignKey("projects.id"), nullable=False, index=True
    )
    user_id = db.Column(
        db.Integer, db.ForeignKey("users.id"), nullable=False, index=True
    )
    is_leader = db.Column(db.Boolean, default=False, nullable=False)
    billing_rate = db.Column(db.Float, nullable=True)

    __table_args__ = (
        db.UniqueConstraint("project_id", "user_id", name="uq_project_user"),
    )


class TimeEntry(db.Model):
    __tablename__ = "time_entries"
    id = db.Column(db.Integer, primary_key=True)
    user_id = db.Column(
        db.Integer, db.ForeignKey("users.id"), nullable=False, index=True
    )
    project_id = db.Column(
        db.Integer, db.ForeignKey("projects.id"), nullable=False, index=True
    )
    description = db.Column(db.Text, nullable=False, default="")
    start_time = db.Column(db.DateTime, nullable=False)
    end_time = db.Column(db.DateTime, nullable=True)
    duration_seconds = db.Column(db.Integer, nullable=True)
    billing_rate = db.Column(db.Float, nullable=False, default=0.0)
    manual = db.Column(db.Boolean, default=False, nullable=False)
    invoiced = db.Column(db.Boolean, default=False, nullable=False, index=True)
    invoice_id = db.Column(
        db.Integer, db.ForeignKey("invoices.id"), nullable=True, index=True
    )
    created_at = db.Column(db.DateTime, default=datetime.utcnow)

    @property
    def is_running(self) -> bool:
        return self.end_time is None

    @property
    def hours(self) -> float:
        if self.duration_seconds is None:
            return 0.0
        return self.duration_seconds / 3600.0

    @property
    def amount(self) -> float:
        return round(self.hours * self.billing_rate, 2)


class Invoice(db.Model):
    __tablename__ = "invoices"
    id = db.Column(db.Integer, primary_key=True)
    project_id = db.Column(
        db.Integer, db.ForeignKey("projects.id"), nullable=False, index=True
    )
    generated_by_id = db.Column(db.Integer, db.ForeignKey("users.id"), nullable=False)
    generated_at = db.Column(db.DateTime, default=datetime.utcnow)
    period_start = db.Column(db.DateTime, nullable=True)
    period_end = db.Column(db.DateTime, nullable=True)
    total_amount = db.Column(db.Float, nullable=False, default=0.0)
    notes = db.Column(db.Text, default="")

    generated_by = db.relationship("User", foreign_keys=[generated_by_id])
    entries = db.relationship("TimeEntry", backref="invoice")
