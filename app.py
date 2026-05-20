import os
from datetime import datetime, timedelta
from functools import wraps

from flask import (
    Flask,
    render_template,
    redirect,
    url_for,
    request,
    flash,
    abort,
    Response,
)
from flask_login import (
    LoginManager,
    login_user,
    logout_user,
    login_required,
    current_user,
)
from sqlalchemy import and_, or_

from models import db, User, Project, ProjectMember, TimeEntry, Invoice


def create_app() -> Flask:
    app = Flask(__name__)
    app.config["SECRET_KEY"] = os.environ.get("SECRET_KEY", "change-me-in-production")
    db_path = os.environ.get(
        "DATABASE_URL", "sqlite:///" + os.path.join(app.instance_path, "timetracker.db")
    )
    os.makedirs(app.instance_path, exist_ok=True)
    app.config["SQLALCHEMY_DATABASE_URI"] = db_path
    app.config["SQLALCHEMY_TRACK_MODIFICATIONS"] = False

    db.init_app(app)
    login_manager = LoginManager(app)
    login_manager.login_view = "login"

    @login_manager.user_loader
    def load_user(user_id: str):
        return db.session.get(User, int(user_id))

    @app.template_filter("dt")
    def fmt_dt(value):
        if value is None:
            return ""
        return value.strftime("%Y-%m-%d %H:%M")

    @app.template_filter("hours")
    def fmt_hours(seconds):
        if seconds is None:
            return "—"
        h = seconds // 3600
        m = (seconds % 3600) // 60
        s = seconds % 60
        return f"{h:d}:{m:02d}:{s:02d}"

    @app.template_filter("money")
    def fmt_money(value):
        if value is None:
            return "$0.00"
        return f"${value:,.2f}"

    # ---------- helpers ----------
    def admin_required(f):
        @wraps(f)
        def wrapper(*args, **kwargs):
            if not current_user.is_authenticated or not current_user.is_admin:
                abort(403)
            return f(*args, **kwargs)

        return wrapper

    def user_can_lead(project: Project) -> bool:
        if current_user.is_admin:
            return True
        membership = ProjectMember.query.filter_by(
            project_id=project.id, user_id=current_user.id, is_leader=True
        ).first()
        return membership is not None

    def projects_for_user(user: User):
        if user.is_admin:
            return Project.query.filter_by(archived=False).order_by(Project.name).all()
        return (
            Project.query.join(ProjectMember)
            .filter(ProjectMember.user_id == user.id, Project.archived == False)
            .order_by(Project.name)
            .all()
        )

    def billing_rate_for(user: User, project: Project) -> float:
        member = ProjectMember.query.filter_by(
            project_id=project.id, user_id=user.id
        ).first()
        if member and member.billing_rate is not None:
            return member.billing_rate
        if project.default_billing_rate:
            return project.default_billing_rate
        return user.default_billing_rate or 0.0

    def running_entry_for(user: User):
        return TimeEntry.query.filter_by(user_id=user.id, end_time=None).first()

    # ---------- auth ----------
    @app.route("/login", methods=["GET", "POST"])
    def login():
        if current_user.is_authenticated:
            return redirect(url_for("dashboard"))
        if request.method == "POST":
            username = request.form.get("username", "").strip()
            password = request.form.get("password", "")
            user = User.query.filter_by(username=username).first()
            if user and user.check_password(password):
                login_user(user)
                return redirect(url_for("dashboard"))
            flash("Invalid username or password.", "error")
        return render_template("login.html")

    @app.route("/logout")
    @login_required
    def logout():
        logout_user()
        return redirect(url_for("login"))

    # ---------- dashboard / timer ----------
    @app.route("/")
    @login_required
    def dashboard():
        running = running_entry_for(current_user)
        projects = projects_for_user(current_user)
        recent = (
            TimeEntry.query.filter_by(user_id=current_user.id)
            .filter(TimeEntry.end_time.isnot(None))
            .order_by(TimeEntry.start_time.desc())
            .limit(10)
            .all()
        )
        # weekly summary
        week_start = datetime.utcnow() - timedelta(days=7)
        weekly = (
            db.session.query(
                Project.name, db.func.sum(TimeEntry.duration_seconds)
            )
            .join(TimeEntry, TimeEntry.project_id == Project.id)
            .filter(
                TimeEntry.user_id == current_user.id,
                TimeEntry.start_time >= week_start,
                TimeEntry.end_time.isnot(None),
            )
            .group_by(Project.name)
            .all()
        )
        return render_template(
            "dashboard.html",
            running=running,
            projects=projects,
            recent=recent,
            weekly=weekly,
        )

    @app.route("/timer/start", methods=["POST"])
    @login_required
    def timer_start():
        if running_entry_for(current_user):
            flash("You already have a timer running. Stop it before starting another.", "error")
            return redirect(url_for("dashboard"))
        project_id = request.form.get("project_id", type=int)
        description = request.form.get("description", "").strip()
        project = db.session.get(Project, project_id)
        if not project:
            flash("Select a project.", "error")
            return redirect(url_for("dashboard"))
        entry = TimeEntry(
            user_id=current_user.id,
            project_id=project.id,
            description=description,
            start_time=datetime.utcnow(),
            billing_rate=billing_rate_for(current_user, project),
            manual=False,
        )
        db.session.add(entry)
        db.session.commit()
        return redirect(url_for("dashboard"))

    @app.route("/timer/stop", methods=["POST"])
    @login_required
    def timer_stop():
        entry = running_entry_for(current_user)
        if not entry:
            flash("No timer is currently running.", "error")
            return redirect(url_for("dashboard"))
        description = request.form.get("description", "").strip()
        if description:
            entry.description = description
        entry.end_time = datetime.utcnow()
        entry.duration_seconds = int(
            (entry.end_time - entry.start_time).total_seconds()
        )
        db.session.commit()
        flash("Time entry saved.", "success")
        return redirect(url_for("dashboard"))

    # ---------- time entries ----------
    @app.route("/entries")
    @login_required
    def entries_list():
        entries = (
            TimeEntry.query.filter_by(user_id=current_user.id)
            .order_by(TimeEntry.start_time.desc())
            .all()
        )
        return render_template("entries.html", entries=entries)

    @app.route("/entries/new", methods=["GET", "POST"])
    @login_required
    def entry_new():
        projects = projects_for_user(current_user)
        if request.method == "POST":
            project_id = request.form.get("project_id", type=int)
            description = request.form.get("description", "").strip()
            start_str = request.form.get("start_time", "").strip()
            end_str = request.form.get("end_time", "").strip()
            project = db.session.get(Project, project_id)
            if not project:
                flash("Select a project.", "error")
                return render_template("entry_form.html", projects=projects, entry=None)
            try:
                start = datetime.strptime(start_str, "%Y-%m-%dT%H:%M")
                end = datetime.strptime(end_str, "%Y-%m-%dT%H:%M")
            except ValueError:
                flash("Invalid date/time format.", "error")
                return render_template("entry_form.html", projects=projects, entry=None)
            if end <= start:
                flash("End time must be after start time.", "error")
                return render_template("entry_form.html", projects=projects, entry=None)
            entry = TimeEntry(
                user_id=current_user.id,
                project_id=project.id,
                description=description,
                start_time=start,
                end_time=end,
                duration_seconds=int((end - start).total_seconds()),
                billing_rate=billing_rate_for(current_user, project),
                manual=True,
            )
            db.session.add(entry)
            db.session.commit()
            flash("Manual time entry added.", "success")
            return redirect(url_for("entries_list"))
        return render_template("entry_form.html", projects=projects, entry=None)

    @app.route("/entries/<int:entry_id>/edit", methods=["GET", "POST"])
    @login_required
    def entry_edit(entry_id: int):
        entry = db.session.get(TimeEntry, entry_id)
        if not entry:
            abort(404)
        if entry.user_id != current_user.id and not current_user.is_admin:
            abort(403)
        if entry.invoiced:
            flash("Cannot edit an entry that has been invoiced.", "error")
            return redirect(url_for("entries_list"))
        projects = projects_for_user(current_user)
        if request.method == "POST":
            project_id = request.form.get("project_id", type=int)
            description = request.form.get("description", "").strip()
            start_str = request.form.get("start_time", "").strip()
            end_str = request.form.get("end_time", "").strip()
            project = db.session.get(Project, project_id)
            if not project:
                flash("Select a project.", "error")
                return render_template("entry_form.html", projects=projects, entry=entry)
            try:
                start = datetime.strptime(start_str, "%Y-%m-%dT%H:%M")
                end = datetime.strptime(end_str, "%Y-%m-%dT%H:%M")
            except ValueError:
                flash("Invalid date/time format.", "error")
                return render_template("entry_form.html", projects=projects, entry=entry)
            if end <= start:
                flash("End time must be after start time.", "error")
                return render_template("entry_form.html", projects=projects, entry=entry)
            entry.project_id = project.id
            entry.description = description
            entry.start_time = start
            entry.end_time = end
            entry.duration_seconds = int((end - start).total_seconds())
            entry.billing_rate = billing_rate_for(entry.user, project)
            db.session.commit()
            flash("Entry updated.", "success")
            return redirect(url_for("entries_list"))
        return render_template("entry_form.html", projects=projects, entry=entry)

    @app.route("/entries/<int:entry_id>/delete", methods=["POST"])
    @login_required
    def entry_delete(entry_id: int):
        entry = db.session.get(TimeEntry, entry_id)
        if not entry:
            abort(404)
        if entry.user_id != current_user.id and not current_user.is_admin:
            abort(403)
        if entry.invoiced:
            flash("Cannot delete an entry that has been invoiced.", "error")
            return redirect(url_for("entries_list"))
        db.session.delete(entry)
        db.session.commit()
        flash("Entry deleted.", "success")
        return redirect(url_for("entries_list"))

    # ---------- projects ----------
    @app.route("/projects")
    @login_required
    def projects_list():
        if current_user.is_admin:
            projects = Project.query.order_by(Project.name).all()
        else:
            projects = projects_for_user(current_user)
        # mark leadership status
        leader_ids = {
            m.project_id
            for m in ProjectMember.query.filter_by(
                user_id=current_user.id, is_leader=True
            ).all()
        }
        return render_template(
            "projects.html", projects=projects, leader_ids=leader_ids
        )

    @app.route("/projects/new", methods=["GET", "POST"])
    @admin_required
    def project_new():
        users = User.query.order_by(User.full_name).all()
        if request.method == "POST":
            name = request.form.get("name", "").strip()
            client_name = request.form.get("client_name", "").strip()
            client_address = request.form.get("client_address", "").strip()
            rate = request.form.get("default_billing_rate", type=float) or 0.0
            if not name or not client_name:
                flash("Project name and client name are required.", "error")
                return render_template("project_form.html", project=None, users=users)
            project = Project(
                name=name,
                client_name=client_name,
                client_address=client_address,
                default_billing_rate=rate,
            )
            db.session.add(project)
            db.session.flush()
            member_ids = request.form.getlist("members", type=int)
            leader_ids = set(request.form.getlist("leaders", type=int))
            for uid in member_ids:
                db.session.add(
                    ProjectMember(
                        project_id=project.id,
                        user_id=uid,
                        is_leader=(uid in leader_ids),
                    )
                )
            db.session.commit()
            flash("Project created.", "success")
            return redirect(url_for("projects_list"))
        return render_template("project_form.html", project=None, users=users)

    @app.route("/projects/<int:project_id>/edit", methods=["GET", "POST"])
    @admin_required
    def project_edit(project_id: int):
        project = db.session.get(Project, project_id)
        if not project:
            abort(404)
        users = User.query.order_by(User.full_name).all()
        if request.method == "POST":
            project.name = request.form.get("name", "").strip()
            project.client_name = request.form.get("client_name", "").strip()
            project.client_address = request.form.get("client_address", "").strip()
            project.default_billing_rate = (
                request.form.get("default_billing_rate", type=float) or 0.0
            )
            project.archived = bool(request.form.get("archived"))
            member_ids = set(request.form.getlist("members", type=int))
            leader_ids = set(request.form.getlist("leaders", type=int))
            # remove members no longer selected
            for m in list(project.members):
                if m.user_id not in member_ids:
                    db.session.delete(m)
            existing = {m.user_id: m for m in project.members if m.user_id in member_ids}
            for uid in member_ids:
                if uid in existing:
                    existing[uid].is_leader = uid in leader_ids
                else:
                    db.session.add(
                        ProjectMember(
                            project_id=project.id,
                            user_id=uid,
                            is_leader=(uid in leader_ids),
                        )
                    )
            db.session.commit()
            flash("Project updated.", "success")
            return redirect(url_for("projects_list"))
        member_ids = {m.user_id for m in project.members}
        leader_ids = {m.user_id for m in project.members if m.is_leader}
        return render_template(
            "project_form.html",
            project=project,
            users=users,
            member_ids=member_ids,
            leader_ids=leader_ids,
        )

    # ---------- reports ----------
    @app.route("/projects/<int:project_id>/report")
    @login_required
    def project_report(project_id: int):
        project = db.session.get(Project, project_id)
        if not project:
            abort(404)
        if not user_can_lead(project):
            abort(403)
        start_str = request.args.get("start", "")
        end_str = request.args.get("end", "")
        q = TimeEntry.query.filter(
            TimeEntry.project_id == project.id,
            TimeEntry.end_time.isnot(None),
        )
        start_dt = end_dt = None
        if start_str:
            try:
                start_dt = datetime.strptime(start_str, "%Y-%m-%d")
                q = q.filter(TimeEntry.start_time >= start_dt)
            except ValueError:
                pass
        if end_str:
            try:
                end_dt = datetime.strptime(end_str, "%Y-%m-%d") + timedelta(days=1)
                q = q.filter(TimeEntry.start_time < end_dt)
            except ValueError:
                pass
        entries = q.order_by(TimeEntry.start_time.asc()).all()
        # summary per user
        per_user = {}
        for e in entries:
            per_user.setdefault(
                e.user.full_name,
                {"seconds": 0, "amount": 0.0, "rate": e.billing_rate},
            )
            per_user[e.user.full_name]["seconds"] += e.duration_seconds or 0
            per_user[e.user.full_name]["amount"] += e.amount
        total_seconds = sum(e.duration_seconds or 0 for e in entries)
        total_amount = sum(e.amount for e in entries)
        return render_template(
            "report.html",
            project=project,
            entries=entries,
            per_user=per_user,
            total_seconds=total_seconds,
            total_amount=total_amount,
            start=start_str,
            end=end_str,
        )

    @app.route("/projects/<int:project_id>/invoice", methods=["GET", "POST"])
    @login_required
    def invoice_new(project_id: int):
        project = db.session.get(Project, project_id)
        if not project:
            abort(404)
        if not user_can_lead(project):
            abort(403)
        if request.method == "POST":
            start_str = request.form.get("start", "")
            end_str = request.form.get("end", "")
            notes = request.form.get("notes", "").strip()
            q = TimeEntry.query.filter(
                TimeEntry.project_id == project.id,
                TimeEntry.end_time.isnot(None),
                TimeEntry.invoiced == False,
            )
            start_dt = end_dt = None
            if start_str:
                try:
                    start_dt = datetime.strptime(start_str, "%Y-%m-%d")
                    q = q.filter(TimeEntry.start_time >= start_dt)
                except ValueError:
                    pass
            if end_str:
                try:
                    end_dt = datetime.strptime(end_str, "%Y-%m-%d") + timedelta(days=1)
                    q = q.filter(TimeEntry.start_time < end_dt)
                except ValueError:
                    pass
            entries = q.order_by(TimeEntry.start_time.asc()).all()
            if not entries:
                flash("No un-invoiced time entries in that period.", "error")
                return redirect(url_for("invoice_new", project_id=project.id))
            total = sum(e.amount for e in entries)
            invoice = Invoice(
                project_id=project.id,
                generated_by_id=current_user.id,
                period_start=start_dt,
                period_end=end_dt,
                total_amount=total,
                notes=notes,
            )
            db.session.add(invoice)
            db.session.flush()
            for e in entries:
                e.invoiced = True
                e.invoice_id = invoice.id
            db.session.commit()
            return redirect(url_for("invoice_view", invoice_id=invoice.id))
        return render_template("invoice_new.html", project=project)

    @app.route("/invoices")
    @login_required
    def invoices_list():
        led_project_ids = [
            m.project_id
            for m in ProjectMember.query.filter_by(
                user_id=current_user.id, is_leader=True
            ).all()
        ]
        if current_user.is_admin:
            invoices = Invoice.query.order_by(Invoice.generated_at.desc()).all()
        else:
            invoices = (
                Invoice.query.filter(Invoice.project_id.in_(led_project_ids))
                .order_by(Invoice.generated_at.desc())
                .all()
            )
        return render_template("invoices.html", invoices=invoices)

    @app.route("/invoices/<int:invoice_id>")
    @login_required
    def invoice_view(invoice_id: int):
        invoice = db.session.get(Invoice, invoice_id)
        if not invoice:
            abort(404)
        if not user_can_lead(invoice.project):
            abort(403)
        entries = sorted(invoice.entries, key=lambda e: e.start_time)
        per_user = {}
        for e in entries:
            per_user.setdefault(
                e.user.full_name,
                {"seconds": 0, "amount": 0.0, "rate": e.billing_rate},
            )
            per_user[e.user.full_name]["seconds"] += e.duration_seconds or 0
            per_user[e.user.full_name]["amount"] += e.amount
        return render_template(
            "invoice.html", invoice=invoice, entries=entries, per_user=per_user
        )

    @app.route("/invoices/<int:invoice_id>/delete", methods=["POST"])
    @login_required
    def invoice_delete(invoice_id: int):
        invoice = db.session.get(Invoice, invoice_id)
        if not invoice:
            abort(404)
        if not user_can_lead(invoice.project):
            abort(403)
        for e in invoice.entries:
            e.invoiced = False
            e.invoice_id = None
        db.session.delete(invoice)
        db.session.commit()
        flash("Invoice deleted; entries returned to un-invoiced pool.", "success")
        return redirect(url_for("invoices_list"))

    # ---------- users (admin) ----------
    @app.route("/users")
    @admin_required
    def users_list():
        users = User.query.order_by(User.full_name).all()
        return render_template("users.html", users=users)

    @app.route("/users/new", methods=["GET", "POST"])
    @admin_required
    def user_new():
        if request.method == "POST":
            username = request.form.get("username", "").strip()
            full_name = request.form.get("full_name", "").strip()
            password = request.form.get("password", "")
            rate = request.form.get("default_billing_rate", type=float) or 0.0
            is_admin = bool(request.form.get("is_admin"))
            if not username or not full_name or not password:
                flash("Username, full name, and password are required.", "error")
                return render_template("user_form.html", user=None)
            if User.query.filter_by(username=username).first():
                flash("Username already taken.", "error")
                return render_template("user_form.html", user=None)
            user = User(
                username=username,
                full_name=full_name,
                default_billing_rate=rate,
                is_admin=is_admin,
            )
            user.set_password(password)
            db.session.add(user)
            db.session.commit()
            flash("User created.", "success")
            return redirect(url_for("users_list"))
        return render_template("user_form.html", user=None)

    @app.route("/users/<int:user_id>/edit", methods=["GET", "POST"])
    @admin_required
    def user_edit(user_id: int):
        user = db.session.get(User, user_id)
        if not user:
            abort(404)
        if request.method == "POST":
            user.full_name = request.form.get("full_name", "").strip()
            user.default_billing_rate = (
                request.form.get("default_billing_rate", type=float) or 0.0
            )
            user.is_admin = bool(request.form.get("is_admin"))
            new_password = request.form.get("password", "")
            if new_password:
                user.set_password(new_password)
            db.session.commit()
            flash("User updated.", "success")
            return redirect(url_for("users_list"))
        return render_template("user_form.html", user=user)

    # ---------- bootstrap ----------
    @app.cli.command("init-db")
    def init_db_cmd():
        """Create tables and seed an initial admin user."""
        db.create_all()
        if not User.query.filter_by(username="admin").first():
            admin = User(
                username="admin",
                full_name="Administrator",
                default_billing_rate=0.0,
                is_admin=True,
            )
            admin.set_password("admin")
            db.session.add(admin)
            db.session.commit()
            print("Created default admin (username=admin, password=admin).")
            print("CHANGE THE PASSWORD IMMEDIATELY via the Users screen.")
        else:
            print("Database already initialized.")

    with app.app_context():
        db.create_all()
        if not User.query.filter_by(username="admin").first():
            admin = User(
                username="admin",
                full_name="Administrator",
                default_billing_rate=0.0,
                is_admin=True,
            )
            admin.set_password("admin")
            db.session.add(admin)
            db.session.commit()

    return app


app = create_app()


if __name__ == "__main__":
    host = os.environ.get("HOST", "0.0.0.0")
    port = int(os.environ.get("PORT", "5000"))
    app.run(host=host, port=port, debug=False)
