from flask import Flask, render_template, request, redirect, url_for, flash, send_file, session, abort, jsonify
import sqlite3, os, json, secrets, hmac, hashlib, shutil, platform
from pathlib import Path
from datetime import datetime, date, timedelta
from io import BytesIO
from functools import wraps
from werkzeug.security import generate_password_hash, check_password_hash
from werkzeug.utils import secure_filename
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.formatting.rule import DataBarRule
from openpyxl.utils import get_column_letter

APP_VERSION = "4.0.0"
BASE_DIR = Path(__file__).resolve().parent.parent
DATA_DIR = BASE_DIR / "data"
DB_PATH = DATA_DIR / "projectplan.db"
SECRET_FILE = DATA_DIR / ".secret_key"

app = Flask(__name__)
app.config.update(
    SESSION_COOKIE_HTTPONLY=True,
    SESSION_COOKIE_SAMESITE="Lax",
    SESSION_COOKIE_SECURE=os.getenv("PROJECT_PLAN_HTTPS", "0") == "1",
    PERMANENT_SESSION_LIFETIME=timedelta(minutes=int(os.getenv("PROJECT_PLAN_SESSION_MINUTES", "60"))),
)

def get_secret():
    DATA_DIR.mkdir(parents=True, exist_ok=True)
    env = os.getenv("PROJECT_PLAN_SECRET")
    if env:
        return env
    if SECRET_FILE.exists():
        return SECRET_FILE.read_text(encoding="utf-8").strip()
    value = secrets.token_hex(32)
    SECRET_FILE.write_text(value, encoding="utf-8")
    return value

app.secret_key = get_secret()

STATUSES = ["Ej startad", "Pågår", "Blockerad", "Klar", "Pausad"]
PRIORITIES = ["Låg", "Normal", "Hög", "Kritisk"]
RISK_STATUSES = ["Öppen", "Bevakas", "Åtgärdas", "Stängd"]
ROLES = ["admin", "pm", "member", "viewer"]

PROJECT_ROLES = ["pm", "member", "viewer"]

LANGUAGES = {
    "sv": {"name": "Svenska", "flag": "🇸🇪"},
    "en": {"name": "English", "flag": "🇬🇧"},
    "de": {"name": "Deutsch", "flag": "🇩🇪"},
    "no": {"name": "Norsk", "flag": "🇳🇴"},
    "da": {"name": "Dansk", "flag": "🇩🇰"},
    "fi": {"name": "Suomi", "flag": "🇫🇮"},
}

TRANSLATIONS = {
    "sv": {
        "dashboard":"Dashboard","admin":"Admin","password":"Lösenord","logout":"Logga ut","api":"API",
        "notifications":"Notiser","pmo":"PMO","resources":"Resurser","my_work":"Mina uppgifter",
        "language":"Språk","projects":"Projekt","new_project":"Nytt projekt","search":"Sök",
        "welcome":"Välkommen tillbaka!","active_projects":"Aktiva projekt","activities":"Aktiviteter",
        "milestones":"Milstolpar","risks":"Risker","project_status":"Projektstatus","upcoming_milestones":"Kommande milstolpar",
        "gantt":"Gantt","kanban":"Kanban","calendar":"Kalender","time_reporting":"Tidrapportering",
        "documents":"Dokument","meetings":"Möten","reports":"Rapporter","settings":"Inställningar",
        "users":"Användare","system":"System","open":"Öppna","save":"Spara","cancel":"Avbryt",
        "login":"Logga in","username":"Användarnamn","display_name":"Visningsnamn","current_password":"Nuvarande lösenord",
        "new_password":"Nytt lösenord","change_password":"Byt lösenord","project_portfolio":"Mina projekt",
        "only_assigned":"Du ser bara projekt där du är medlem.","admin_all_projects":"Administratörsvy: alla projekt.",
        "brand_tagline":"PLAN · GENOMFÖR · LYCKAS"
    },
    "en": {
        "dashboard":"Dashboard","admin":"Admin","password":"Password","logout":"Sign out","api":"API",
        "notifications":"Notifications","pmo":"PMO","resources":"Resources","my_work":"My Work",
        "language":"Language","projects":"Projects","new_project":"New project","search":"Search",
        "welcome":"Welcome back!","active_projects":"Active projects","activities":"Activities",
        "milestones":"Milestones","risks":"Risks","project_status":"Project status","upcoming_milestones":"Upcoming milestones",
        "gantt":"Gantt","kanban":"Kanban","calendar":"Calendar","time_reporting":"Time reporting",
        "documents":"Documents","meetings":"Meetings","reports":"Reports","settings":"Settings",
        "users":"Users","system":"System","open":"Open","save":"Save","cancel":"Cancel",
        "login":"Sign in","username":"Username","display_name":"Display name","current_password":"Current password",
        "new_password":"New password","change_password":"Change password","project_portfolio":"My projects",
        "only_assigned":"You only see projects where you are a member.","admin_all_projects":"Administrator view: all projects.",
        "brand_tagline":"PLAN · EXECUTE · SUCCEED"
    },
    "de": {
        "dashboard":"Dashboard","admin":"Admin","password":"Passwort","logout":"Abmelden","api":"API",
        "notifications":"Benachrichtigungen","pmo":"PMO","resources":"Ressourcen","my_work":"Meine Aufgaben",
        "language":"Sprache","projects":"Projekte","new_project":"Neues Projekt","search":"Suchen",
        "welcome":"Willkommen zurück!","active_projects":"Aktive Projekte","activities":"Aufgaben",
        "milestones":"Meilensteine","risks":"Risiken","project_status":"Projektstatus","upcoming_milestones":"Kommende Meilensteine",
        "gantt":"Gantt","kanban":"Kanban","calendar":"Kalender","time_reporting":"Zeiterfassung",
        "documents":"Dokumente","meetings":"Meetings","reports":"Berichte","settings":"Einstellungen",
        "users":"Benutzer","system":"System","open":"Öffnen","save":"Speichern","cancel":"Abbrechen",
        "login":"Anmelden","username":"Benutzername","display_name":"Anzeigename","current_password":"Aktuelles Passwort",
        "new_password":"Neues Passwort","change_password":"Passwort ändern","project_portfolio":"Meine Projekte",
        "only_assigned":"Sie sehen nur Projekte, denen Sie zugewiesen sind.","admin_all_projects":"Administratoransicht: alle Projekte.",
        "brand_tagline":"PLANEN · UMSETZEN · ERFOLG"
    },
    "no": {
        "dashboard":"Dashboard","admin":"Admin","password":"Passord","logout":"Logg ut","api":"API",
        "notifications":"Varsler","pmo":"PMO","resources":"Ressurser","my_work":"Mine oppgaver",
        "language":"Språk","projects":"Prosjekter","new_project":"Nytt prosjekt","search":"Søk",
        "welcome":"Velkommen tilbake!","active_projects":"Aktive prosjekter","activities":"Aktiviteter",
        "milestones":"Milepæler","risks":"Risikoer","project_status":"Prosjektstatus","upcoming_milestones":"Kommende milepæler",
        "gantt":"Gantt","kanban":"Kanban","calendar":"Kalender","time_reporting":"Timeregistrering",
        "documents":"Dokumenter","meetings":"Møter","reports":"Rapporter","settings":"Innstillinger",
        "users":"Brukere","system":"System","open":"Åpne","save":"Lagre","cancel":"Avbryt",
        "login":"Logg inn","username":"Brukernavn","display_name":"Visningsnavn","current_password":"Nåværende passord",
        "new_password":"Nytt passord","change_password":"Endre passord","project_portfolio":"Mine prosjekter",
        "only_assigned":"Du ser bare prosjekter du er medlem av.","admin_all_projects":"Administratorvisning: alle prosjekter.",
        "brand_tagline":"PLANLEGG · GJENNOMFØR · LYKKES"
    },
    "da": {
        "dashboard":"Dashboard","admin":"Admin","password":"Adgangskode","logout":"Log ud","api":"API",
        "notifications":"Notifikationer","pmo":"PMO","resources":"Ressourcer","my_work":"Mine opgaver",
        "language":"Sprog","projects":"Projekter","new_project":"Nyt projekt","search":"Søg",
        "welcome":"Velkommen tilbage!","active_projects":"Aktive projekter","activities":"Aktiviteter",
        "milestones":"Milepæle","risks":"Risici","project_status":"Projektstatus","upcoming_milestones":"Kommende milepæle",
        "gantt":"Gantt","kanban":"Kanban","calendar":"Kalender","time_reporting":"Tidsregistrering",
        "documents":"Dokumenter","meetings":"Møder","reports":"Rapporter","settings":"Indstillinger",
        "users":"Brugere","system":"System","open":"Åbn","save":"Gem","cancel":"Annuller",
        "login":"Log ind","username":"Brugernavn","display_name":"Visningsnavn","current_password":"Nuværende adgangskode",
        "new_password":"Ny adgangskode","change_password":"Skift adgangskode","project_portfolio":"Mine projekter",
        "only_assigned":"Du ser kun projekter, hvor du er medlem.","admin_all_projects":"Administratorvisning: alle projekter.",
        "brand_tagline":"PLANLÆG · UDFØR · LYKKES"
    },
    "fi": {
        "dashboard":"Kojelauta","admin":"Ylläpito","password":"Salasana","logout":"Kirjaudu ulos","api":"API",
        "notifications":"Ilmoitukset","pmo":"PMO","resources":"Resurssit","my_work":"Omat tehtävät",
        "language":"Kieli","projects":"Projektit","new_project":"Uusi projekti","search":"Haku",
        "welcome":"Tervetuloa takaisin!","active_projects":"Aktiiviset projektit","activities":"Tehtävät",
        "milestones":"Välitavoitteet","risks":"Riskit","project_status":"Projektin tila","upcoming_milestones":"Tulevat välitavoitteet",
        "gantt":"Gantt","kanban":"Kanban","calendar":"Kalenteri","time_reporting":"Työajanseuranta",
        "documents":"Dokumentit","meetings":"Kokoukset","reports":"Raportit","settings":"Asetukset",
        "users":"Käyttäjät","system":"Järjestelmä","open":"Avaa","save":"Tallenna","cancel":"Peruuta",
        "login":"Kirjaudu sisään","username":"Käyttäjänimi","display_name":"Näyttönimi","current_password":"Nykyinen salasana",
        "new_password":"Uusi salasana","change_password":"Vaihda salasana","project_portfolio":"Omat projektit",
        "only_assigned":"Näet vain projektit, joissa olet jäsenenä.","admin_all_projects":"Ylläpitäjän näkymä: kaikki projektit.",
        "brand_tagline":"SUUNNITTELE · TOTEUTA · ONNISTU"
    },
}


def db():
    DATA_DIR.mkdir(parents=True, exist_ok=True)
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA foreign_keys=ON")
    return conn

def ensure_column(conn, table, column, definition):
    cols = {r["name"] for r in conn.execute(f"PRAGMA table_info({table})")}
    if column not in cols:
        conn.execute(f"ALTER TABLE {table} ADD COLUMN {column} {definition}")

def init_db():
    with db() as conn:
        conn.executescript("""
        CREATE TABLE IF NOT EXISTS schema_migrations (
            version TEXT PRIMARY KEY,
            applied_at TEXT NOT NULL
        );

        CREATE TABLE IF NOT EXISTS users (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            username TEXT NOT NULL UNIQUE,
            display_name TEXT NOT NULL DEFAULT '',
            password_hash TEXT NOT NULL,
            role TEXT NOT NULL DEFAULT 'member',
            active INTEGER NOT NULL DEFAULT 1,
            force_password_change INTEGER NOT NULL DEFAULT 0,
            failed_logins INTEGER NOT NULL DEFAULT 0,
            locked_until TEXT DEFAULT '',
            last_login TEXT DEFAULT '',
            created_at TEXT NOT NULL
        );

        CREATE TABLE IF NOT EXISTS projects (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            name TEXT NOT NULL,
            customer TEXT DEFAULT '',
            project_manager TEXT DEFAULT '',
            description TEXT DEFAULT '',
            start_date TEXT DEFAULT '',
            end_date TEXT DEFAULT '',
            template_name TEXT DEFAULT '',
            created_by INTEGER,
            created_at TEXT NOT NULL,
            FOREIGN KEY(created_by) REFERENCES users(id) ON DELETE SET NULL
        );

        CREATE TABLE IF NOT EXISTS project_members (
            project_id INTEGER NOT NULL,
            user_id INTEGER NOT NULL,
            project_role TEXT NOT NULL DEFAULT 'viewer',
            added_at TEXT NOT NULL DEFAULT '',
            added_by INTEGER,
            PRIMARY KEY(project_id, user_id),
            FOREIGN KEY(project_id) REFERENCES projects(id) ON DELETE CASCADE,
            FOREIGN KEY(user_id) REFERENCES users(id) ON DELETE CASCADE,
            FOREIGN KEY(added_by) REFERENCES users(id) ON DELETE SET NULL
        );

        CREATE TABLE IF NOT EXISTS tasks (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            project_id INTEGER NOT NULL,
            wbs TEXT DEFAULT '',
            title TEXT NOT NULL,
            owner TEXT DEFAULT '',
            start_date TEXT DEFAULT '',
            end_date TEXT DEFAULT '',
            actual_start TEXT DEFAULT '',
            actual_end TEXT DEFAULT '',
            status TEXT DEFAULT 'Ej startad',
            priority TEXT DEFAULT 'Normal',
            progress INTEGER DEFAULT 0,
            milestone INTEGER DEFAULT 0,
            dependencies TEXT DEFAULT '',
            notes TEXT DEFAULT '',
            sort_order INTEGER DEFAULT 0,
            FOREIGN KEY(project_id) REFERENCES projects(id) ON DELETE CASCADE
        );

        CREATE TABLE IF NOT EXISTS task_dependencies (
            task_id INTEGER NOT NULL,
            depends_on_task_id INTEGER NOT NULL,
            PRIMARY KEY(task_id, depends_on_task_id),
            FOREIGN KEY(task_id) REFERENCES tasks(id) ON DELETE CASCADE,
            FOREIGN KEY(depends_on_task_id) REFERENCES tasks(id) ON DELETE CASCADE
        );

        CREATE TABLE IF NOT EXISTS risks (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            project_id INTEGER NOT NULL,
            kind TEXT NOT NULL DEFAULT 'Risk',
            title TEXT NOT NULL,
            description TEXT DEFAULT '',
            probability INTEGER DEFAULT 3,
            impact INTEGER DEFAULT 3,
            owner TEXT DEFAULT '',
            action TEXT DEFAULT '',
            status TEXT DEFAULT 'Öppen',
            due_date TEXT DEFAULT '',
            created_at TEXT NOT NULL,
            FOREIGN KEY(project_id) REFERENCES projects(id) ON DELETE CASCADE
        );

        CREATE TABLE IF NOT EXISTS comments (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            project_id INTEGER NOT NULL,
            task_id INTEGER,
            user_id INTEGER,
            body TEXT NOT NULL,
            created_at TEXT NOT NULL,
            FOREIGN KEY(project_id) REFERENCES projects(id) ON DELETE CASCADE,
            FOREIGN KEY(task_id) REFERENCES tasks(id) ON DELETE CASCADE,
            FOREIGN KEY(user_id) REFERENCES users(id) ON DELETE SET NULL
        );

        CREATE TABLE IF NOT EXISTS audit_log (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            project_id INTEGER,
            user_id INTEGER,
            entity_type TEXT NOT NULL,
            entity_id INTEGER,
            action TEXT NOT NULL,
            details TEXT DEFAULT '',
            created_at TEXT NOT NULL,
            ip_address TEXT DEFAULT '',
            FOREIGN KEY(project_id) REFERENCES projects(id) ON DELETE CASCADE,
            FOREIGN KEY(user_id) REFERENCES users(id) ON DELETE SET NULL
        );

        CREATE TABLE IF NOT EXISTS baselines (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            project_id INTEGER NOT NULL,
            name TEXT NOT NULL,
            created_at TEXT NOT NULL,
            created_by INTEGER,
            snapshot_json TEXT NOT NULL,
            FOREIGN KEY(project_id) REFERENCES projects(id) ON DELETE CASCADE,
            FOREIGN KEY(created_by) REFERENCES users(id) ON DELETE SET NULL
        );

        CREATE TABLE IF NOT EXISTS project_templates (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            name TEXT NOT NULL UNIQUE,
            description TEXT DEFAULT '',
            tasks_json TEXT NOT NULL,
            created_at TEXT NOT NULL
        );

        CREATE TABLE IF NOT EXISTS system_settings (
            key TEXT PRIMARY KEY,
            value TEXT NOT NULL
        );

        CREATE TABLE IF NOT EXISTS api_keys (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            user_id INTEGER NOT NULL,
            name TEXT NOT NULL,
            key_hash TEXT NOT NULL UNIQUE,
            prefix TEXT NOT NULL,
            active INTEGER NOT NULL DEFAULT 1,
            created_at TEXT NOT NULL,
            last_used TEXT DEFAULT '',
            FOREIGN KEY(user_id) REFERENCES users(id) ON DELETE CASCADE
        );
        CREATE TABLE IF NOT EXISTS webhooks (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            project_id INTEGER,
            name TEXT NOT NULL,
            target_url TEXT NOT NULL,
            event_type TEXT NOT NULL DEFAULT 'project.updated',
            active INTEGER NOT NULL DEFAULT 1
        );
        CREATE TABLE IF NOT EXISTS custom_field_definitions (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            entity_type TEXT NOT NULL,
            name TEXT NOT NULL,
            field_type TEXT NOT NULL DEFAULT 'text',
            options_json TEXT NOT NULL DEFAULT '[]',
            required INTEGER NOT NULL DEFAULT 0
        );
        CREATE TABLE IF NOT EXISTS custom_field_values (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            field_id INTEGER NOT NULL,
            entity_type TEXT NOT NULL,
            entity_id INTEGER NOT NULL,
            value TEXT DEFAULT '',
            UNIQUE(field_id,entity_type,entity_id),
            FOREIGN KEY(field_id) REFERENCES custom_field_definitions(id) ON DELETE CASCADE
        );
        CREATE TABLE IF NOT EXISTS dashboard_preferences (
            user_id INTEGER PRIMARY KEY,
            layout_json TEXT NOT NULL DEFAULT '{}',
            FOREIGN KEY(user_id) REFERENCES users(id) ON DELETE CASCADE
        );

        CREATE TABLE IF NOT EXISTS documents (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            project_id INTEGER NOT NULL,
            title TEXT NOT NULL,
            body TEXT NOT NULL DEFAULT '',
            updated_by INTEGER,
            updated_at TEXT NOT NULL,
            FOREIGN KEY(project_id) REFERENCES projects(id) ON DELETE CASCADE
        );
        CREATE TABLE IF NOT EXISTS attachments (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            project_id INTEGER NOT NULL,
            entity_type TEXT NOT NULL DEFAULT 'project',
            entity_id INTEGER,
            original_name TEXT NOT NULL,
            stored_name TEXT NOT NULL,
            uploaded_by INTEGER,
            uploaded_at TEXT NOT NULL,
            FOREIGN KEY(project_id) REFERENCES projects(id) ON DELETE CASCADE
        );
        CREATE TABLE IF NOT EXISTS meetings (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            project_id INTEGER NOT NULL,
            title TEXT NOT NULL,
            meeting_date TEXT NOT NULL,
            attendees TEXT DEFAULT '',
            notes TEXT DEFAULT '',
            FOREIGN KEY(project_id) REFERENCES projects(id) ON DELETE CASCADE
        );
        CREATE TABLE IF NOT EXISTS action_items (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            project_id INTEGER NOT NULL,
            meeting_id INTEGER,
            title TEXT NOT NULL,
            owner TEXT DEFAULT '',
            due_date TEXT DEFAULT '',
            status TEXT NOT NULL DEFAULT 'Open',
            FOREIGN KEY(project_id) REFERENCES projects(id) ON DELETE CASCADE,
            FOREIGN KEY(meeting_id) REFERENCES meetings(id) ON DELETE SET NULL
        );
        CREATE TABLE IF NOT EXISTS notifications (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            user_id INTEGER NOT NULL,
            project_id INTEGER,
            message TEXT NOT NULL,
            link TEXT DEFAULT '',
            is_read INTEGER NOT NULL DEFAULT 0,
            created_at TEXT NOT NULL,
            FOREIGN KEY(user_id) REFERENCES users(id) ON DELETE CASCADE
        );
        CREATE TABLE IF NOT EXISTS automation_rules (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            project_id INTEGER,
            name TEXT NOT NULL,
            trigger_type TEXT NOT NULL,
            action_type TEXT NOT NULL,
            config_json TEXT NOT NULL DEFAULT '{}',
            active INTEGER NOT NULL DEFAULT 1
        );

        CREATE TABLE IF NOT EXISTS portfolios (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            name TEXT NOT NULL,
            description TEXT DEFAULT '',
            owner TEXT DEFAULT ''
        );
        CREATE TABLE IF NOT EXISTS portfolio_projects (
            portfolio_id INTEGER NOT NULL,
            project_id INTEGER NOT NULL,
            PRIMARY KEY(portfolio_id,project_id),
            FOREIGN KEY(portfolio_id) REFERENCES portfolios(id) ON DELETE CASCADE,
            FOREIGN KEY(project_id) REFERENCES projects(id) ON DELETE CASCADE
        );
        CREATE TABLE IF NOT EXISTS status_reports (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            project_id INTEGER NOT NULL,
            report_date TEXT NOT NULL,
            overall_rag TEXT NOT NULL DEFAULT 'Green',
            scope_rag TEXT NOT NULL DEFAULT 'Green',
            schedule_rag TEXT NOT NULL DEFAULT 'Green',
            budget_rag TEXT NOT NULL DEFAULT 'Green',
            resources_rag TEXT NOT NULL DEFAULT 'Green',
            summary TEXT DEFAULT '',
            achievements TEXT DEFAULT '',
            next_steps TEXT DEFAULT '',
            created_by INTEGER,
            created_at TEXT NOT NULL,
            FOREIGN KEY(project_id) REFERENCES projects(id) ON DELETE CASCADE
        );
        CREATE TABLE IF NOT EXISTS raid_items (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            project_id INTEGER NOT NULL,
            item_type TEXT NOT NULL,
            title TEXT NOT NULL,
            owner TEXT DEFAULT '',
            status TEXT DEFAULT 'Open',
            due_date TEXT DEFAULT '',
            details TEXT DEFAULT '',
            FOREIGN KEY(project_id) REFERENCES projects(id) ON DELETE CASCADE
        );
        CREATE TABLE IF NOT EXISTS decisions (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            project_id INTEGER NOT NULL,
            title TEXT NOT NULL,
            decision TEXT NOT NULL,
            decided_by TEXT DEFAULT '',
            decision_date TEXT NOT NULL,
            FOREIGN KEY(project_id) REFERENCES projects(id) ON DELETE CASCADE
        );
        CREATE TABLE IF NOT EXISTS change_requests (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            project_id INTEGER NOT NULL,
            title TEXT NOT NULL,
            description TEXT DEFAULT '',
            impact_scope TEXT DEFAULT '',
            impact_days INTEGER NOT NULL DEFAULT 0,
            impact_cost REAL NOT NULL DEFAULT 0,
            status TEXT NOT NULL DEFAULT 'Proposed',
            requested_by INTEGER,
            created_at TEXT NOT NULL,
            FOREIGN KEY(project_id) REFERENCES projects(id) ON DELETE CASCADE
        );
        CREATE TABLE IF NOT EXISTS approvals (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            project_id INTEGER NOT NULL,
            entity_type TEXT NOT NULL,
            entity_id INTEGER NOT NULL,
            requested_by INTEGER,
            approved_by INTEGER,
            status TEXT NOT NULL DEFAULT 'Pending',
            comment TEXT DEFAULT '',
            created_at TEXT NOT NULL,
            decided_at TEXT DEFAULT '',
            FOREIGN KEY(project_id) REFERENCES projects(id) ON DELETE CASCADE
        );

        CREATE TABLE IF NOT EXISTS resource_profiles (
            user_id INTEGER PRIMARY KEY,
            weekly_capacity REAL NOT NULL DEFAULT 40,
            hourly_cost REAL NOT NULL DEFAULT 0,
            hourly_rate REAL NOT NULL DEFAULT 0,
            FOREIGN KEY(user_id) REFERENCES users(id) ON DELETE CASCADE
        );

        CREATE TABLE IF NOT EXISTS time_entries (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            project_id INTEGER NOT NULL,
            task_id INTEGER,
            user_id INTEGER NOT NULL,
            work_date TEXT NOT NULL,
            hours REAL NOT NULL,
            note TEXT DEFAULT '',
            created_at TEXT NOT NULL,
            FOREIGN KEY(project_id) REFERENCES projects(id) ON DELETE CASCADE,
            FOREIGN KEY(task_id) REFERENCES tasks(id) ON DELETE SET NULL,
            FOREIGN KEY(user_id) REFERENCES users(id) ON DELETE CASCADE
        );

        CREATE TABLE IF NOT EXISTS project_costs (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            project_id INTEGER NOT NULL,
            category TEXT NOT NULL DEFAULT 'External',
            description TEXT NOT NULL,
            planned REAL NOT NULL DEFAULT 0,
            actual REAL NOT NULL DEFAULT 0,
            cost_date TEXT DEFAULT '',
            FOREIGN KEY(project_id) REFERENCES projects(id) ON DELETE CASCADE
        );

        CREATE TABLE IF NOT EXISTS project_finance (
            project_id INTEGER PRIMARY KEY,
            budget REAL NOT NULL DEFAULT 0,
            currency TEXT NOT NULL DEFAULT 'SEK',
            FOREIGN KEY(project_id) REFERENCES projects(id) ON DELETE CASCADE
        );

        CREATE TABLE IF NOT EXISTS task_links (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            project_id INTEGER NOT NULL,
            predecessor_id INTEGER NOT NULL,
            successor_id INTEGER NOT NULL,
            link_type TEXT NOT NULL DEFAULT 'FS',
            lag_days INTEGER NOT NULL DEFAULT 0,
            FOREIGN KEY(project_id) REFERENCES projects(id) ON DELETE CASCADE,
            FOREIGN KEY(predecessor_id) REFERENCES tasks(id) ON DELETE CASCADE,
            FOREIGN KEY(successor_id) REFERENCES tasks(id) ON DELETE CASCADE
        );

        CREATE TABLE IF NOT EXISTS saved_views (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            user_id INTEGER NOT NULL,
            project_id INTEGER,
            name TEXT NOT NULL,
            view_type TEXT NOT NULL DEFAULT 'list',
            filter_json TEXT NOT NULL DEFAULT '{}',
            created_at TEXT NOT NULL,
            FOREIGN KEY(user_id) REFERENCES users(id) ON DELETE CASCADE,
            FOREIGN KEY(project_id) REFERENCES projects(id) ON DELETE CASCADE
        );
        """)

        conn.executescript("""

        CREATE TABLE IF NOT EXISTS resource_allocations (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            project_id INTEGER NOT NULL,
            user_id INTEGER,
            resource_name TEXT NOT NULL,
            week_start TEXT NOT NULL,
            allocation_pct INTEGER NOT NULL DEFAULT 0,
            planned_hours REAL NOT NULL DEFAULT 0,
            FOREIGN KEY(project_id) REFERENCES projects(id) ON DELETE CASCADE,
            FOREIGN KEY(user_id) REFERENCES users(id) ON DELETE SET NULL
        );

        """)

        conn.executescript("""

        CREATE TABLE IF NOT EXISTS programs (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            name TEXT NOT NULL UNIQUE,
            owner TEXT NOT NULL DEFAULT '',
            description TEXT NOT NULL DEFAULT ''
        );
        CREATE TABLE IF NOT EXISTS program_projects (
            program_id INTEGER NOT NULL,
            project_id INTEGER NOT NULL,
            PRIMARY KEY(program_id,project_id),
            FOREIGN KEY(program_id) REFERENCES programs(id) ON DELETE CASCADE,
            FOREIGN KEY(project_id) REFERENCES projects(id) ON DELETE CASCADE
        );

        """)

        conn.executescript("""

        CREATE TABLE IF NOT EXISTS document_versions (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            document_id INTEGER NOT NULL,
            version_no INTEGER NOT NULL,
            content TEXT NOT NULL DEFAULT '',
            created_by INTEGER,
            created_at TEXT NOT NULL,
            FOREIGN KEY(document_id) REFERENCES documents(id) ON DELETE CASCADE,
            FOREIGN KEY(created_by) REFERENCES users(id) ON DELETE SET NULL
        );
        CREATE TABLE IF NOT EXISTS mentions (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            project_id INTEGER NOT NULL,
            mentioned_user_id INTEGER NOT NULL,
            source_type TEXT NOT NULL,
            source_id INTEGER,
            message TEXT NOT NULL,
            created_at TEXT NOT NULL,
            read_at TEXT NOT NULL DEFAULT '',
            FOREIGN KEY(project_id) REFERENCES projects(id) ON DELETE CASCADE,
            FOREIGN KEY(mentioned_user_id) REFERENCES users(id) ON DELETE CASCADE
        );

        """)

        conn.executescript("""

        CREATE TABLE IF NOT EXISTS report_definitions (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            name TEXT NOT NULL,
            report_type TEXT NOT NULL DEFAULT 'portfolio',
            config_json TEXT NOT NULL DEFAULT '{}',
            owner_user_id INTEGER,
            created_at TEXT NOT NULL,
            FOREIGN KEY(owner_user_id) REFERENCES users(id) ON DELETE SET NULL
        );
        CREATE TABLE IF NOT EXISTS oidc_settings (
            id INTEGER PRIMARY KEY CHECK (id=1),
            enabled INTEGER NOT NULL DEFAULT 0,
            issuer TEXT NOT NULL DEFAULT '',
            client_id TEXT NOT NULL DEFAULT '',
            scopes TEXT NOT NULL DEFAULT 'openid profile email'
        );
        INSERT OR IGNORE INTO oidc_settings(id) VALUES(1);

        """)

        conn.executescript("""

        """)

        conn.executescript("""
CREATE TABLE IF NOT EXISTS favorites (
    user_id INTEGER NOT NULL,
    entity_type TEXT NOT NULL,
    entity_id INTEGER NOT NULL,
    created_at TEXT NOT NULL,
    PRIMARY KEY(user_id,entity_type,entity_id),
    FOREIGN KEY(user_id) REFERENCES users(id) ON DELETE CASCADE
);
CREATE TABLE IF NOT EXISTS recent_items (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    user_id INTEGER NOT NULL,
    entity_type TEXT NOT NULL,
    entity_id INTEGER NOT NULL,
    label TEXT NOT NULL,
    url TEXT NOT NULL,
    viewed_at TEXT NOT NULL,
    FOREIGN KEY(user_id) REFERENCES users(id) ON DELETE CASCADE
);
        """)

        conn.executescript("""
CREATE TABLE IF NOT EXISTS task_constraints (
    task_id INTEGER PRIMARY KEY,
    constraint_type TEXT NOT NULL DEFAULT 'ASAP',
    constraint_date TEXT DEFAULT '',
    calendar_name TEXT DEFAULT 'Standard',
    FOREIGN KEY(task_id) REFERENCES tasks(id) ON DELETE CASCADE
);
CREATE TABLE IF NOT EXISTS planning_scenarios (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    project_id INTEGER NOT NULL,
    name TEXT NOT NULL,
    snapshot_json TEXT NOT NULL,
    created_by INTEGER,
    created_at TEXT NOT NULL,
    FOREIGN KEY(project_id) REFERENCES projects(id) ON DELETE CASCADE
);
        """)

        conn.executescript("""
CREATE TABLE IF NOT EXISTS automation_executions (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    rule_id INTEGER,
    project_id INTEGER,
    status TEXT NOT NULL,
    details TEXT DEFAULT '',
    started_at TEXT NOT NULL,
    finished_at TEXT DEFAULT '',
    FOREIGN KEY(rule_id) REFERENCES automation_rules(id) ON DELETE SET NULL
);
        """)

        conn.executescript("""
CREATE TABLE IF NOT EXISTS scheduled_reports (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    project_id INTEGER,
    name TEXT NOT NULL,
    format TEXT NOT NULL DEFAULT 'pdf',
    cadence TEXT NOT NULL DEFAULT 'weekly',
    recipient TEXT DEFAULT '',
    active INTEGER NOT NULL DEFAULT 1,
    created_by INTEGER,
    created_at TEXT NOT NULL
);
        """)

        conn.executescript("""
CREATE TABLE IF NOT EXISTS integration_settings (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    provider TEXT NOT NULL UNIQUE,
    enabled INTEGER NOT NULL DEFAULT 0,
    base_url TEXT DEFAULT '',
    config_json TEXT NOT NULL DEFAULT '{}',
    updated_at TEXT NOT NULL
);
CREATE TABLE IF NOT EXISTS webhook_deliveries (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    webhook_id INTEGER,
    event_type TEXT NOT NULL,
    payload_json TEXT NOT NULL,
    status TEXT NOT NULL DEFAULT 'Pending',
    attempts INTEGER NOT NULL DEFAULT 0,
    last_error TEXT DEFAULT '',
    created_at TEXT NOT NULL,
    FOREIGN KEY(webhook_id) REFERENCES webhooks(id) ON DELETE CASCADE
);
        """)

        conn.executescript("""
CREATE TABLE IF NOT EXISTS user_sessions (
    id TEXT PRIMARY KEY,
    user_id INTEGER NOT NULL,
    created_at TEXT NOT NULL,
    last_seen TEXT NOT NULL,
    ip_address TEXT DEFAULT '',
    user_agent TEXT DEFAULT '',
    revoked_at TEXT DEFAULT '',
    FOREIGN KEY(user_id) REFERENCES users(id) ON DELETE CASCADE
);
CREATE TABLE IF NOT EXISTS login_history (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    user_id INTEGER,
    username TEXT NOT NULL,
    success INTEGER NOT NULL,
    ip_address TEXT DEFAULT '',
    user_agent TEXT DEFAULT '',
    created_at TEXT NOT NULL,
    FOREIGN KEY(user_id) REFERENCES users(id) ON DELETE SET NULL
);
CREATE TABLE IF NOT EXISTS security_events (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    severity TEXT NOT NULL,
    event_type TEXT NOT NULL,
    user_id INTEGER,
    details TEXT DEFAULT '',
    ip_address TEXT DEFAULT '',
    created_at TEXT NOT NULL
);
        """)

        conn.executescript("""
CREATE TABLE IF NOT EXISTS project_environments (
    id INTEGER PRIMARY KEY AUTOINCREMENT, project_id INTEGER NOT NULL, name TEXT NOT NULL, base_url TEXT DEFAULT '', owner TEXT DEFAULT '', status TEXT DEFAULT 'Ready', notes TEXT DEFAULT '', UNIQUE(project_id,name), FOREIGN KEY(project_id) REFERENCES projects(id) ON DELETE CASCADE);
CREATE TABLE IF NOT EXISTS deliverables (
    id INTEGER PRIMARY KEY AUTOINCREMENT, project_id INTEGER NOT NULL, name TEXT NOT NULL, owner TEXT DEFAULT '', due_date TEXT DEFAULT '', status TEXT DEFAULT 'Planned', acceptance_criteria TEXT DEFAULT '', FOREIGN KEY(project_id) REFERENCES projects(id) ON DELETE CASCADE);
CREATE TABLE IF NOT EXISTS interface_register (
    id INTEGER PRIMARY KEY AUTOINCREMENT, project_id INTEGER NOT NULL, name TEXT NOT NULL, source_system TEXT DEFAULT '', target_system TEXT DEFAULT '', protocol TEXT DEFAULT '', owner TEXT DEFAULT '', status TEXT DEFAULT 'Design', notes TEXT DEFAULT '', FOREIGN KEY(project_id) REFERENCES projects(id) ON DELETE CASCADE);
CREATE TABLE IF NOT EXISTS test_cycles (
    id INTEGER PRIMARY KEY AUTOINCREMENT, project_id INTEGER NOT NULL, name TEXT NOT NULL, test_type TEXT NOT NULL DEFAULT 'SIT', start_date TEXT DEFAULT '', end_date TEXT DEFAULT '', status TEXT DEFAULT 'Planned', passed INTEGER NOT NULL DEFAULT 0, failed INTEGER NOT NULL DEFAULT 0, FOREIGN KEY(project_id) REFERENCES projects(id) ON DELETE CASCADE);
CREATE TABLE IF NOT EXISTS requirements_traceability (
    id INTEGER PRIMARY KEY AUTOINCREMENT, project_id INTEGER NOT NULL, requirement_id TEXT NOT NULL, requirement TEXT NOT NULL, deliverable_id INTEGER, test_cycle_id INTEGER, status TEXT DEFAULT 'Open', UNIQUE(project_id,requirement_id), FOREIGN KEY(project_id) REFERENCES projects(id) ON DELETE CASCADE);
CREATE TABLE IF NOT EXISTS cutover_items (
    id INTEGER PRIMARY KEY AUTOINCREMENT, project_id INTEGER NOT NULL, sequence_no INTEGER NOT NULL DEFAULT 0, title TEXT NOT NULL, owner TEXT DEFAULT '', planned_at TEXT DEFAULT '', status TEXT DEFAULT 'Planned', rollback_step TEXT DEFAULT '', FOREIGN KEY(project_id) REFERENCES projects(id) ON DELETE CASCADE);
        """)

        conn.executescript("""
CREATE TABLE IF NOT EXISTS project_health_snapshots (
    id INTEGER PRIMARY KEY AUTOINCREMENT, project_id INTEGER NOT NULL, snapshot_date TEXT NOT NULL, progress INTEGER NOT NULL, overdue_count INTEGER NOT NULL, blocked_count INTEGER NOT NULL, high_risk_count INTEGER NOT NULL, open_change_count INTEGER NOT NULL, health_score INTEGER NOT NULL, details_json TEXT NOT NULL DEFAULT '{}', UNIQUE(project_id,snapshot_date), FOREIGN KEY(project_id) REFERENCES projects(id) ON DELETE CASCADE);
CREATE TABLE IF NOT EXISTS assistant_queries (
    id INTEGER PRIMARY KEY AUTOINCREMENT, user_id INTEGER NOT NULL, project_id INTEGER, question TEXT NOT NULL, answer TEXT NOT NULL, created_at TEXT NOT NULL, FOREIGN KEY(user_id) REFERENCES users(id) ON DELETE CASCADE);
        """)

        # Migration from older versions
        ensure_column(conn, "users", "force_password_change", "INTEGER NOT NULL DEFAULT 0")
        ensure_column(conn, "users", "failed_logins", "INTEGER NOT NULL DEFAULT 0")
        ensure_column(conn, "users", "locked_until", "TEXT DEFAULT ''")
        ensure_column(conn, "users", "last_login", "TEXT DEFAULT ''")
        ensure_column(conn, "projects", "created_by", "INTEGER")
        ensure_column(conn, "project_members", "added_at", "TEXT NOT NULL DEFAULT ''")
        ensure_column(conn, "project_members", "added_by", "INTEGER")
        ensure_column(conn, "audit_log", "ip_address", "TEXT DEFAULT ''")
        ensure_column(conn, "tasks", "planned_hours", "REAL DEFAULT 0")
        ensure_column(conn, "tasks", "remaining_hours", "REAL DEFAULT 0")
        ensure_column(conn, "tasks", "parent_task_id", "INTEGER")
        ensure_column(conn, "tasks", "duration_days", "INTEGER DEFAULT 1")
        ensure_column(conn, "tasks", "owner_user_id", "INTEGER")

        ensure_column(conn, "change_requests", "reason", "TEXT DEFAULT ''")
        ensure_column(conn, "change_requests", "decided_by", "INTEGER")
        ensure_column(conn, "change_requests", "decided_at", "TEXT DEFAULT ''")
        ensure_column(conn, "decisions", "owner", "TEXT DEFAULT ''")
        ensure_column(conn, "decisions", "decided_at", "TEXT DEFAULT ''")

        ensure_column(conn, "api_keys", "scopes", "TEXT NOT NULL DEFAULT 'read'")
        ensure_column(conn, "api_keys", "expires_at", "TEXT DEFAULT ''")

        ensure_column(conn, "projects", "archived_at", "TEXT DEFAULT ''")
        ensure_column(conn, "projects", "deleted_at", "TEXT DEFAULT ''")
        ensure_column(conn, "tasks", "deleted_at", "TEXT DEFAULT ''")

        conn.execute("INSERT OR IGNORE INTO schema_migrations(version,applied_at) VALUES(?,?)",
                     (APP_VERSION, datetime.now().isoformat(timespec="seconds")))
        conn.execute("INSERT OR IGNORE INTO system_settings(key,value) VALUES('allow_project_creation','pm')")
        conn.execute("INSERT OR IGNORE INTO system_settings(key,value) VALUES('lockout_attempts','5')")
        conn.execute("INSERT OR IGNORE INTO system_settings(key,value) VALUES('lockout_minutes','15')")

        templates = {
            "Standardprojekt":[
                {"wbs":"1","title":"Initiering","milestone":0},
                {"wbs":"1.1","title":"Projektmål och scope","milestone":0},
                {"wbs":"1.2","title":"Kickoff","milestone":1},
                {"wbs":"2","title":"Planering","milestone":0},
                {"wbs":"2.1","title":"Detaljplan","milestone":0},
                {"wbs":"2.2","title":"Riskworkshop","milestone":0},
                {"wbs":"3","title":"Genomförande","milestone":0},
                {"wbs":"4","title":"Verifiering och acceptans","milestone":0},
                {"wbs":"5","title":"Överlämning","milestone":1}
            ],
            "IT-/integrationsprojekt":[
                {"wbs":"1","title":"Förstudie och krav","milestone":0},
                {"wbs":"2","title":"Arkitektur och design","milestone":1},
                {"wbs":"3","title":"Utveckling/konfiguration","milestone":0},
                {"wbs":"4","title":"Integrationstest","milestone":0},
                {"wbs":"5","title":"UAT","milestone":1},
                {"wbs":"6","title":"Driftsättning","milestone":1},
                {"wbs":"7","title":"Hypercare och avslut","milestone":0}
            ],
            "LIMS-projekt":[
                {"wbs":"1","title":"Process- och kravanalys","milestone":0},
                {"wbs":"2","title":"LIMS-design","milestone":1},
                {"wbs":"3","title":"Konfiguration och utveckling","milestone":0},
                {"wbs":"4","title":"Instrument/integrationer","milestone":0},
                {"wbs":"5","title":"Verifiering/validering","milestone":0},
                {"wbs":"6","title":"Utbildning","milestone":0},
                {"wbs":"7","title":"Go-live","milestone":1},
                {"wbs":"8","title":"Stabilisering","milestone":0}
            ]
        }
        for name,tasks in templates.items():
            conn.execute("""INSERT OR IGNORE INTO project_templates(name,description,tasks_json,created_at)
                            VALUES(?,?,?,?)""",
                         (name,f"Inbyggd mall: {name}",json.dumps(tasks,ensure_ascii=False),
                          datetime.now().isoformat(timespec="seconds")))

def setting(key, default=""):
    with db() as conn:
        row = conn.execute("SELECT value FROM system_settings WHERE key=?", (key,)).fetchone()
    return row["value"] if row else default

def user_count():
    with db() as conn:
        return conn.execute("SELECT COUNT(*) c FROM users").fetchone()["c"]

def current_user():
    uid = session.get("user_id")
    if not uid: return None
    with db() as conn:
        return conn.execute("SELECT * FROM users WHERE id=? AND active=1",(uid,)).fetchone()

def csrf_token():
    token = session.get("_csrf")
    if not token:
        token = secrets.token_urlsafe(32)
        session["_csrf"] = token
    return token


def active_language():
    lang = session.get("lang", "sv")
    return lang if lang in LANGUAGES else "sv"

def tr(key):
    lang = active_language()
    return TRANSLATIONS.get(lang, TRANSLATIONS["sv"]).get(key, TRANSLATIONS["sv"].get(key, key))

@app.get("/language/<lang>")
def set_language(lang):
    if lang not in LANGUAGES:
        abort(404)
    session["lang"] = lang
    return redirect(request.referrer or url_for("index"))

@app.context_processor
def inject_i18n():
    return dict(t=tr, active_lang=active_language(), languages=LANGUAGES)

@app.context_processor
def inject_globals():
    return dict(current_user=current_user(), app_version=APP_VERSION, csrf_token=csrf_token())

@app.before_request
def bootstrap_and_security():
    init_db()
    if request.method in ("POST","PUT","PATCH","DELETE"):
        sent = request.form.get("_csrf") or request.headers.get("X-CSRF-Token","")
        expected = session.get("_csrf","")
        if not expected or not sent or not hmac.compare_digest(sent, expected):
            abort(400, "Ogiltig eller saknad CSRF-token.")
    if session.get("user_id"):
        session.permanent = True
        u = current_user()
        if u and u["force_password_change"] and request.endpoint not in ("change_password","logout","static"):
            return redirect(url_for("change_password"))

def login_required(fn):
    @wraps(fn)
    def wrapper(*args,**kwargs):
        if user_count()==0: return redirect(url_for("setup"))
        if not current_user(): return redirect(url_for("login",next=request.path))
        return fn(*args,**kwargs)
    return wrapper

def role_required(*roles):
    def deco(fn):
        @wraps(fn)
        def wrapper(*args,**kwargs):
            u=current_user()
            if not u or u["role"] not in roles: abort(403)
            return fn(*args,**kwargs)
        return wrapper
    return deco

def project_access(project_id, write=False, manager=False):
    u=current_user()
    if not u: abort(401)
    if u["role"]=="admin": return "admin"
    with db() as conn:
        row=conn.execute("""SELECT pm.project_role FROM project_members pm
                            JOIN projects p ON p.id=pm.project_id
                            WHERE pm.project_id=? AND pm.user_id=?""",(project_id,u["id"])).fetchone()
    if not row: abort(403)
    pr=row["project_role"]
    if manager and pr!="pm": abort(403)
    if write and pr not in ("pm","member"): abort(403)
    return pr

def audit(project_id, entity_type, entity_id, action, details=""):
    u=current_user()
    ip=request.headers.get("X-Forwarded-For",request.remote_addr or "").split(",")[0].strip()
    with db() as conn:
        conn.execute("""INSERT INTO audit_log(project_id,user_id,entity_type,entity_id,action,details,created_at,ip_address)
                        VALUES(?,?,?,?,?,?,?,?)""",
                     (project_id,u["id"] if u else None,entity_type,entity_id,action,details,
                      datetime.now().isoformat(timespec="seconds"),ip))

def project_or_404(project_id, write=False, manager=False):
    project_access(project_id,write=write,manager=manager)
    with db() as conn:
        row=conn.execute("SELECT * FROM projects WHERE id=?",(project_id,)).fetchone()
    if not row: abort(404)
    return row

def task_or_404(project_id,task_id,write=False):
    project_access(project_id,write=write)
    with db() as conn:
        row=conn.execute("SELECT * FROM tasks WHERE id=? AND project_id=?",(task_id,project_id)).fetchone()
    if not row: abort(404)
    return row

def parse_date(s):
    if not s:return None
    try:return datetime.strptime(s,"%Y-%m-%d").date()
    except:return None

def derived_status(task):
    if task["status"]=="Klar" or int(task["progress"] or 0)>=100:return "Klar"
    today=date.today(); end=parse_date(task["end_date"]); start=parse_date(task["start_date"])
    if task["status"]=="Blockerad":return "Blockerad"
    if end and end<today:return "Försenad"
    if start and start>today:return "Kommande"
    if int(task["progress"] or 0)>0 or (start and start<=today):return "Pågår"
    return task["status"] or "Ej startad"

def task_health(task):
    return {"Försenad":"danger","Blockerad":"warning","Klar":"success","Pågår":"info"}.get(derived_status(task),"muted")

@app.route("/setup",methods=["GET","POST"])
def setup():
    if user_count()>0:return redirect(url_for("login"))
    if request.method=="POST":
        username=request.form["username"].strip()
        password=request.form["password"]
        if len(password)<10:
            flash("Lösenordet måste vara minst 10 tecken.","danger")
            return render_template("setup.html")
        with db() as conn:
            conn.execute("""INSERT INTO users(username,display_name,password_hash,role,active,force_password_change,created_at)
                            VALUES(?,?,?,?,1,0,?)""",
                         (username,request.form.get("display_name","").strip() or username,
                          generate_password_hash(password),"admin",datetime.now().isoformat(timespec="seconds")))
        flash("Administratör skapad. Logga in.","success")
        return redirect(url_for("login"))
    return render_template("setup.html")

@app.route("/login",methods=["GET","POST"])
def login():
    if user_count()==0:return redirect(url_for("setup"))
    if request.method=="POST":
        username=request.form["username"].strip()
        with db() as conn:
            u=conn.execute("SELECT * FROM users WHERE username=?",(username,)).fetchone()
            if u:
                locked_until=None
                if u["locked_until"]:
                    try: locked_until=datetime.fromisoformat(u["locked_until"])
                    except: pass
                if not u["active"]:
                    flash("Kontot är inaktiverat.","danger"); return render_template("login.html")
                if locked_until and locked_until>datetime.now():
                    flash(f"Kontot är tillfälligt låst till {locked_until:%H:%M}.","danger")
                    return render_template("login.html")
                if check_password_hash(u["password_hash"],request.form["password"]):
                    conn.execute("UPDATE users SET failed_logins=0,locked_until='',last_login=? WHERE id=?",
                                 (datetime.now().isoformat(timespec="seconds"),u["id"]))
                    session.clear(); session["user_id"]=u["id"]; session["_csrf"]=secrets.token_urlsafe(32); session.permanent=True
                    return redirect(request.args.get("next") or url_for("index"))
                attempts=int(u["failed_logins"] or 0)+1
                limit=int(setting("lockout_attempts","5"))
                lock_until=""
                if attempts>=limit:
                    lock_until=(datetime.now()+timedelta(minutes=int(setting("lockout_minutes","15")))).isoformat(timespec="seconds")
                    attempts=0
                conn.execute("UPDATE users SET failed_logins=?,locked_until=? WHERE id=?",(attempts,lock_until,u["id"]))
        flash("Fel användarnamn eller lösenord.","danger")
    return render_template("login.html")

@app.get("/logout")
def logout():
    session.clear()
    return redirect(url_for("login"))

@app.route("/change-password",methods=["GET","POST"])
@login_required
def change_password():
    u=current_user()
    if request.method=="POST":
        new=request.form["new_password"]
        if len(new)<10:
            flash("Det nya lösenordet måste vara minst 10 tecken.","danger")
        else:
            with db() as conn:
                dbu=conn.execute("SELECT * FROM users WHERE id=?",(u["id"],)).fetchone()
                if not check_password_hash(dbu["password_hash"],request.form["current_password"]):
                    flash("Nuvarande lösenord är fel.","danger")
                else:
                    conn.execute("UPDATE users SET password_hash=?,force_password_change=0 WHERE id=?",
                                 (generate_password_hash(new),u["id"]))
                    flash("Lösenordet är ändrat.","success")
                    return redirect(url_for("index"))
    return render_template("change_password.html")

@app.get("/health")
def health():
    try:
        with db() as conn: conn.execute("SELECT 1").fetchone()
        return jsonify(status="ok",version=APP_VERSION,database="ok"),200
    except Exception as e:
        return jsonify(status="error",version=APP_VERSION,error=str(e)),503

@app.route("/")
@login_required
def index():
    u=current_user()
    with db() as conn:
        if u["role"]=="admin":
            projects=conn.execute("""SELECT p.*,COUNT(t.id) task_count,COALESCE(ROUND(AVG(t.progress)),0) avg_progress,
                SUM(CASE WHEN t.status='Klar' OR t.progress>=100 THEN 1 ELSE 0 END) done_count,
                SUM(CASE WHEN t.status='Blockerad' THEN 1 ELSE 0 END) blocked_count
                FROM projects p LEFT JOIN tasks t ON t.project_id=p.id GROUP BY p.id ORDER BY p.id DESC""").fetchall()
            upcoming=conn.execute("""SELECT t.*,p.name project_name FROM tasks t JOIN projects p ON p.id=t.project_id
                WHERE t.end_date<>'' AND t.progress<100 ORDER BY t.end_date LIMIT 12""").fetchall()
            open_risks=conn.execute("""SELECT r.*,p.name project_name FROM risks r JOIN projects p ON p.id=r.project_id
                WHERE r.status<>'Stängd' ORDER BY (r.probability*r.impact) DESC LIMIT 10""").fetchall()
        else:
            projects=conn.execute("""SELECT p.*,COUNT(t.id) task_count,COALESCE(ROUND(AVG(t.progress)),0) avg_progress,
                SUM(CASE WHEN t.status='Klar' OR t.progress>=100 THEN 1 ELSE 0 END) done_count,
                SUM(CASE WHEN t.status='Blockerad' THEN 1 ELSE 0 END) blocked_count
                FROM projects p JOIN project_members pm ON pm.project_id=p.id AND pm.user_id=?
                LEFT JOIN tasks t ON t.project_id=p.id GROUP BY p.id ORDER BY p.id DESC""",(u["id"],)).fetchall()
            upcoming=conn.execute("""SELECT t.*,p.name project_name FROM tasks t JOIN projects p ON p.id=t.project_id
                JOIN project_members pm ON pm.project_id=p.id AND pm.user_id=?
                WHERE t.end_date<>'' AND t.progress<100 ORDER BY t.end_date LIMIT 12""",(u["id"],)).fetchall()
            open_risks=conn.execute("""SELECT r.*,p.name project_name FROM risks r JOIN projects p ON p.id=r.project_id
                JOIN project_members pm ON pm.project_id=p.id AND pm.user_id=?
                WHERE r.status<>'Stängd' ORDER BY (r.probability*r.impact) DESC LIMIT 10""",(u["id"],)).fetchall()
    overdue=sum(1 for t in upcoming if parse_date(t["end_date"]) and parse_date(t["end_date"])<date.today())
    return render_template("index.html",projects=projects,upcoming=upcoming,open_risks=open_risks,overdue=overdue)

@app.route("/projects/new",methods=["GET","POST"])
@login_required
@role_required("admin","pm")
def new_project():
    u=current_user()
    with db() as conn: templates=conn.execute("SELECT * FROM project_templates ORDER BY name").fetchall()
    if request.method=="POST":
        with db() as conn:
            cur=conn.execute("""INSERT INTO projects(name,customer,project_manager,description,start_date,end_date,template_name,created_by,created_at)
                                VALUES(?,?,?,?,?,?,?,?,?)""",
                             (request.form["name"].strip(),request.form.get("customer","").strip(),
                              request.form.get("project_manager","").strip(),request.form.get("description","").strip(),
                              request.form.get("start_date",""),request.form.get("end_date",""),
                              request.form.get("template_name",""),u["id"],datetime.now().isoformat(timespec="seconds")))
            pid=cur.lastrowid
            conn.execute("""INSERT INTO project_members(project_id,user_id,project_role,added_at,added_by)
                            VALUES(?,?,?,?,?)""",(pid,u["id"],"pm",datetime.now().isoformat(timespec="seconds"),u["id"]))
            tmpl=request.form.get("template_name","")
            if tmpl:
                row=conn.execute("SELECT tasks_json FROM project_templates WHERE name=?",(tmpl,)).fetchone()
                if row:
                    for i,item in enumerate(json.loads(row["tasks_json"])):
                        conn.execute("""INSERT INTO tasks(project_id,wbs,title,status,priority,progress,milestone,sort_order)
                                        VALUES(?,?,?,?,?,?,?,?)""",(pid,item.get("wbs",""),item["title"],"Ej startad","Normal",0,item.get("milestone",0),i))
        audit(pid,"project",pid,"created","Projekt skapat")
        return redirect(url_for("project",project_id=pid))
    return render_template("project_form.html",project=None,templates=templates)

@app.route("/projects/<int:project_id>")
@login_required
def project(project_id):
    p=project_or_404(project_id)
    q=request.args.get("q","").strip(); status=request.args.get("status",""); owner=request.args.get("owner",""); priority=request.args.get("priority",""); sort=request.args.get("sort","wbs")
    sql="SELECT * FROM tasks WHERE project_id=?"; args=[project_id]
    if q:
        like=f"%{q}%"; sql+=" AND (title LIKE ? OR wbs LIKE ? OR owner LIKE ? OR notes LIKE ?)"; args += [like,like,like,like]
    if status: sql+=" AND status=?"; args.append(status)
    if owner: sql+=" AND owner=?"; args.append(owner)
    if priority: sql+=" AND priority=?"; args.append(priority)
    allowed={"wbs":"wbs COLLATE NOCASE","end":"end_date","start":"start_date","priority":"priority","owner":"owner COLLATE NOCASE"}
    sql+=" ORDER BY "+allowed.get(sort,allowed["wbs"])+",sort_order,id"
    with db() as conn:
        tasks=conn.execute(sql,args).fetchall()
        all_tasks=conn.execute("SELECT * FROM tasks WHERE project_id=? ORDER BY wbs,id",(project_id,)).fetchall()
        risks=conn.execute("SELECT * FROM risks WHERE project_id=? ORDER BY (probability*impact) DESC,id DESC",(project_id,)).fetchall()
        comments=conn.execute("""SELECT c.*,u.display_name FROM comments c LEFT JOIN users u ON u.id=c.user_id
                                 WHERE c.project_id=? ORDER BY c.id DESC LIMIT 20""",(project_id,)).fetchall()
        history=conn.execute("""SELECT a.*,u.display_name FROM audit_log a LEFT JOIN users u ON u.id=a.user_id
                                WHERE a.project_id=? ORDER BY a.id DESC LIMIT 40""",(project_id,)).fetchall()
        baselines=conn.execute("SELECT * FROM baselines WHERE project_id=? ORDER BY id DESC",(project_id,)).fetchall()
        owners=conn.execute("SELECT DISTINCT owner FROM tasks WHERE project_id=? AND owner<>'' ORDER BY owner",(project_id,)).fetchall()
        members=conn.execute("""SELECT pm.*,u.username,u.display_name,u.role global_role FROM project_members pm
                                JOIN users u ON u.id=pm.user_id WHERE pm.project_id=? ORDER BY pm.project_role,u.display_name""",(project_id,)).fetchall()
        candidates=conn.execute("""SELECT u.* FROM users u WHERE u.active=1 AND u.id NOT IN
                                   (SELECT user_id FROM project_members WHERE project_id=?)
                                   ORDER BY u.display_name""",(project_id,)).fetchall()
    avg=round(sum(int(t["progress"] or 0) for t in all_tasks)/len(all_tasks)) if all_tasks else 0
    enriched=[dict(t,derived_status=derived_status(t),health=task_health(t)) for t in tasks]
    u=current_user()
    project_role="admin" if u["role"]=="admin" else next((m["project_role"] for m in members if m["user_id"]==u["id"]),"viewer")
    return render_template("project.html",project=p,tasks=enriched,all_tasks=all_tasks,risks=risks,comments=comments,
                           history=history,baselines=baselines,owners=owners,members=members,candidates=candidates,
                           project_role=project_role,statuses=STATUSES,priorities=PRIORITIES,risk_statuses=RISK_STATUSES,
                           avg=avg,done=sum(1 for t in all_tasks if derived_status(t)=="Klar"),
                           blocked=sum(1 for t in all_tasks if derived_status(t)=="Blockerad"),
                           overdue=sum(1 for t in all_tasks if derived_status(t)=="Försenad"),
                           upcoming=sum(1 for t in all_tasks if parse_date(t["end_date"]) and date.today()<=parse_date(t["end_date"])<=date.today()+timedelta(days=7) and int(t["progress"] or 0)<100))

@app.route("/projects/<int:project_id>/edit",methods=["GET","POST"])
@login_required
def edit_project(project_id):
    p=project_or_404(project_id,manager=True)
    with db() as conn: templates=conn.execute("SELECT * FROM project_templates ORDER BY name").fetchall()
    if request.method=="POST":
        with db() as conn:
            conn.execute("""UPDATE projects SET name=?,customer=?,project_manager=?,description=?,start_date=?,end_date=? WHERE id=?""",
                         (request.form["name"].strip(),request.form.get("customer","").strip(),
                          request.form.get("project_manager","").strip(),request.form.get("description","").strip(),
                          request.form.get("start_date",""),request.form.get("end_date",""),project_id))
        audit(project_id,"project",project_id,"updated","Projektinformation uppdaterad")
        return redirect(url_for("project",project_id=project_id))
    return render_template("project_form.html",project=p,templates=templates)

@app.post("/projects/<int:project_id>/members/add")
@login_required
def add_project_member(project_id):
    project_or_404(project_id,manager=True)
    uid=int(request.form["user_id"]); prole=request.form.get("project_role","viewer")
    if prole not in PROJECT_ROLES: prole="viewer"
    u=current_user()
    with db() as conn:
        conn.execute("""INSERT OR REPLACE INTO project_members(project_id,user_id,project_role,added_at,added_by)
                        VALUES(?,?,?,?,?)""",(project_id,uid,prole,datetime.now().isoformat(timespec="seconds"),u["id"]))
    audit(project_id,"member",uid,"added",prole)
    return redirect(url_for("project",project_id=project_id)+"#members")

@app.post("/projects/<int:project_id>/members/<int:user_id>/role")
@login_required
def change_project_member_role(project_id,user_id):
    project_or_404(project_id,manager=True)
    prole=request.form.get("project_role","viewer")
    if prole not in PROJECT_ROLES: abort(400)
    with db() as conn:
        conn.execute("UPDATE project_members SET project_role=? WHERE project_id=? AND user_id=?",(prole,project_id,user_id))
    audit(project_id,"member",user_id,"role-changed",prole)
    return redirect(url_for("project",project_id=project_id)+"#members")

@app.post("/projects/<int:project_id>/members/<int:user_id>/remove")
@login_required
def remove_project_member(project_id,user_id):
    project_or_404(project_id,manager=True)
    u=current_user()
    if u["id"]==user_id and u["role"]!="admin":
        flash("Du kan inte ta bort dig själv som projektledare.","danger")
    else:
        with db() as conn:
            conn.execute("DELETE FROM project_members WHERE project_id=? AND user_id=?",(project_id,user_id))
        audit(project_id,"member",user_id,"removed","")
    return redirect(url_for("project",project_id=project_id)+"#members")

@app.route("/projects/<int:project_id>/tasks/new",methods=["GET","POST"])
@login_required
def new_task(project_id):
    p=project_or_404(project_id,write=True)
    with db() as conn: candidates=conn.execute("SELECT * FROM tasks WHERE project_id=? ORDER BY wbs,id",(project_id,)).fetchall()
    if request.method=="POST":
        with db() as conn:
            cur=conn.execute("""INSERT INTO tasks(project_id,wbs,title,owner,start_date,end_date,actual_start,actual_end,status,priority,progress,milestone,notes,sort_order)
                                VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?)""",
                             (project_id,request.form.get("wbs","").strip(),request.form["title"].strip(),request.form.get("owner","").strip(),
                              request.form.get("start_date",""),request.form.get("end_date",""),request.form.get("actual_start",""),request.form.get("actual_end",""),
                              request.form.get("status","Ej startad"),request.form.get("priority","Normal"),
                              max(0,min(100,int(request.form.get("progress","0") or 0))),1 if request.form.get("milestone")=="on" else 0,
                              request.form.get("notes","").strip(),int(request.form.get("sort_order","0") or 0)))
            tid=cur.lastrowid
            for dep in request.form.getlist("dependency_ids"):
                if dep.isdigit() and int(dep)!=tid:
                    conn.execute("INSERT OR IGNORE INTO task_dependencies(task_id,depends_on_task_id) VALUES(?,?)",(tid,int(dep)))
        audit(project_id,"task",tid,"created",request.form["title"].strip())
        return redirect(url_for("project",project_id=project_id))
    return render_template("task_form.html",project=p,task=None,candidates=candidates,selected_deps=[],statuses=STATUSES,priorities=PRIORITIES)

@app.route("/projects/<int:project_id>/tasks/<int:task_id>/edit",methods=["GET","POST"])
@login_required
def edit_task(project_id,task_id):
    p=project_or_404(project_id,write=True); t=task_or_404(project_id,task_id,write=True)
    with db() as conn:
        candidates=conn.execute("SELECT * FROM tasks WHERE project_id=? AND id<>? ORDER BY wbs,id",(project_id,task_id)).fetchall()
        selected=[r["depends_on_task_id"] for r in conn.execute("SELECT depends_on_task_id FROM task_dependencies WHERE task_id=?",(task_id,))]
    if request.method=="POST":
        prog=max(0,min(100,int(request.form.get("progress","0") or 0))); stat=request.form.get("status","Ej startad")
        if stat=="Klar": prog=100
        with db() as conn:
            conn.execute("""UPDATE tasks SET wbs=?,title=?,owner=?,start_date=?,end_date=?,actual_start=?,actual_end=?,status=?,priority=?,progress=?,milestone=?,notes=?,sort_order=?
                            WHERE id=? AND project_id=?""",
                         (request.form.get("wbs","").strip(),request.form["title"].strip(),request.form.get("owner","").strip(),
                          request.form.get("start_date",""),request.form.get("end_date",""),request.form.get("actual_start",""),request.form.get("actual_end",""),
                          stat,request.form.get("priority","Normal"),prog,1 if request.form.get("milestone")=="on" else 0,
                          request.form.get("notes","").strip(),int(request.form.get("sort_order","0") or 0),task_id,project_id))
            conn.execute("DELETE FROM task_dependencies WHERE task_id=?",(task_id,))
            for dep in request.form.getlist("dependency_ids"):
                if dep.isdigit() and int(dep)!=task_id:
                    conn.execute("INSERT OR IGNORE INTO task_dependencies(task_id,depends_on_task_id) VALUES(?,?)",(task_id,int(dep)))
        audit(project_id,"task",task_id,"updated",request.form["title"].strip())
        return redirect(url_for("project",project_id=project_id))
    return render_template("task_form.html",project=p,task=t,candidates=candidates,selected_deps=selected,statuses=STATUSES,priorities=PRIORITIES)

@app.post("/projects/<int:project_id>/tasks/<int:task_id>/delete")
@login_required
def delete_task(project_id,task_id):
    project_or_404(project_id,manager=True)
    t=task_or_404(project_id,task_id)
    with db() as conn: conn.execute("DELETE FROM tasks WHERE id=? AND project_id=?",(task_id,project_id))
    audit(project_id,"task",task_id,"deleted",t["title"])
    return redirect(url_for("project",project_id=project_id))

@app.post("/projects/<int:project_id>/risks/new")
@login_required
def new_risk(project_id):
    project_or_404(project_id,write=True)
    with db() as conn:
        cur=conn.execute("""INSERT INTO risks(project_id,kind,title,description,probability,impact,owner,action,status,due_date,created_at)
                            VALUES(?,?,?,?,?,?,?,?,?,?,?)""",
                         (project_id,request.form.get("kind","Risk"),request.form["title"].strip(),request.form.get("description","").strip(),
                          int(request.form.get("probability","3")),int(request.form.get("impact","3")),request.form.get("owner","").strip(),
                          request.form.get("action","").strip(),request.form.get("status","Öppen"),request.form.get("due_date",""),
                          datetime.now().isoformat(timespec="seconds")))
        rid=cur.lastrowid
    audit(project_id,"risk",rid,"created",request.form["title"].strip())
    return redirect(url_for("project",project_id=project_id)+"#risks")

@app.post("/projects/<int:project_id>/risks/<int:risk_id>/delete")
@login_required
def delete_risk(project_id,risk_id):
    project_or_404(project_id,manager=True)
    with db() as conn: conn.execute("DELETE FROM risks WHERE id=? AND project_id=?",(risk_id,project_id))
    audit(project_id,"risk",risk_id,"deleted","")
    return redirect(url_for("project",project_id=project_id)+"#risks")

@app.post("/projects/<int:project_id>/comments/new")
@login_required
def new_comment(project_id):
    project_or_404(project_id)
    body=request.form.get("body","").strip()
    if body:
        u=current_user()
        with db() as conn: conn.execute("INSERT INTO comments(project_id,user_id,body,created_at) VALUES(?,?,?,?)",
                                        (project_id,u["id"],body,datetime.now().isoformat(timespec="seconds")))
        audit(project_id,"comment",None,"created",body[:120])
    return redirect(url_for("project",project_id=project_id)+"#comments")

@app.post("/projects/<int:project_id>/baseline")
@login_required
def create_baseline(project_id):
    p=project_or_404(project_id,manager=True)
    with db() as conn:
        tasks=[dict(r) for r in conn.execute("SELECT * FROM tasks WHERE project_id=? ORDER BY wbs,id",(project_id,))]
        snap=json.dumps({"project":dict(p),"tasks":tasks},ensure_ascii=False)
        u=current_user()
        conn.execute("INSERT INTO baselines(project_id,name,created_at,created_by,snapshot_json) VALUES(?,?,?,?,?)",
                     (project_id,request.form.get("name","").strip() or f"Baseline {datetime.now():%Y-%m-%d %H:%M}",
                      datetime.now().isoformat(timespec="seconds"),u["id"],snap))
    audit(project_id,"baseline",None,"created","Ny baseline")
    return redirect(url_for("project",project_id=project_id)+"#baseline")

@app.route("/projects/<int:project_id>/gantt")
@login_required
def gantt(project_id):
    p=project_or_404(project_id)
    with db() as conn:
        tasks=conn.execute("SELECT * FROM tasks WHERE project_id=? AND start_date<>'' AND end_date<>'' ORDER BY start_date,wbs,id",(project_id,)).fetchall()
        deps=conn.execute("""SELECT d.task_id,d.depends_on_task_id,t.title dep_title FROM task_dependencies d
                             JOIN tasks t ON t.id=d.depends_on_task_id WHERE d.task_id IN
                             (SELECT id FROM tasks WHERE project_id=?)""",(project_id,)).fetchall()
    if not tasks:return render_template("gantt.html",project=p,rows=[],months=[],range_start=None,range_end=None,deps={})
    starts=[parse_date(t["start_date"]) for t in tasks]; ends=[parse_date(t["end_date"]) for t in tasks]
    rs=min(starts); re=max(ends); total=max((re-rs).days+1,1); rows=[]
    for t in tasks:
        s=parse_date(t["start_date"]); e=parse_date(t["end_date"])
        rows.append(dict(t,left=((s-rs).days/total)*100,width=max(((e-s).days+1)/total*100,.7),derived_status=derived_status(t)))
    months=[]; cur=rs.replace(day=1)
    while cur<=re:
        nextm=(cur.replace(day=28)+timedelta(days=4)).replace(day=1); seg_start=max(cur,rs); seg_end=min(nextm-timedelta(days=1),re)
        months.append({"label":cur.strftime("%Y-%m"),"left":((seg_start-rs).days/total)*100,"width":((seg_end-seg_start).days+1)/total*100}); cur=nextm
    depmap={}
    for d in deps:depmap.setdefault(d["task_id"],[]).append(d["dep_title"])
    return render_template("gantt.html",project=p,rows=rows,months=months,range_start=rs,range_end=re,deps=depmap)

def autosize(ws):
    for col_cells in ws.columns:
        max_len=0; col=get_column_letter(col_cells[0].column)
        for cell in col_cells:max_len=max(max_len,len("" if cell.value is None else str(cell.value)))
        ws.column_dimensions[col].width=min(max(max_len+2,10),45)

def style_header(ws,row=1):
    fill=PatternFill("solid",fgColor="1F4E78"); font=Font(color="FFFFFF",bold=True)
    for c in ws[row]:c.fill=fill;c.font=font;c.alignment=Alignment(vertical="center")

def build_workbook(project,tasks,risks,baselines,members):
    wb=Workbook(); ws=wb.active; ws.title="Projektplan"
    headers=["WBS","Aktivitet","Ansvarig","Plan start","Plan slut","Faktisk start","Faktiskt slut","Status","Automatisk status","Prioritet","Progress %","Milstolpe","Kommentar"]
    ws.append(headers); style_header(ws)
    for t in tasks:ws.append([t["wbs"],t["title"],t["owner"],t["start_date"],t["end_date"],t["actual_start"],t["actual_end"],t["status"],derived_status(t),t["priority"],t["progress"],"Ja" if t["milestone"] else "Nej",t["notes"]])
    ws.freeze_panes="A2"; ws.auto_filter.ref=ws.dimensions
    if ws.max_row>1:ws.conditional_formatting.add(f"K2:K{ws.max_row}",DataBarRule(start_type="num",start_value=0,end_type="num",end_value=100))
    autosize(ws)
    sm=wb.create_sheet("Sammanfattning"); 
    for row in [["Projekt",project["name"]],["Kund",project["customer"]],["Projektledare",project["project_manager"]],["Plan start",project["start_date"]],["Plan slut",project["end_date"]],["Version",APP_VERSION]]:
        sm.append(row)
    avg=round(sum(int(t["progress"] or 0) for t in tasks)/len(tasks)) if tasks else 0
    sm.append(["Total progress",avg]); sm.append(["Aktiviteter",len(tasks)]); autosize(sm)
    ms=wb.create_sheet("Milstolpar"); ms.append(["WBS","Milstolpe","Planerat datum","Status","Progress"]); style_header(ms)
    for t in tasks:
        if t["milestone"]:ms.append([t["wbs"],t["title"],t["end_date"],derived_status(t),t["progress"]])
    autosize(ms)
    rw=wb.create_sheet("Risker & Issues"); rw.append(["Typ","Titel","Sannolikhet","Konsekvens","Riskvärde","Ansvarig","Åtgärd","Status","Förfallodatum"]); style_header(rw)
    for r in risks:rw.append([r["kind"],r["title"],r["probability"],r["impact"],r["probability"]*r["impact"],r["owner"],r["action"],r["status"],r["due_date"]])
    autosize(rw)
    mw=wb.create_sheet("Projektmedlemmar"); mw.append(["Namn","Användare","Projektroll","Global roll"]); style_header(mw)
    for m in members:mw.append([m["display_name"],m["username"],m["project_role"],m["global_role"]])
    autosize(mw)
    return wb

@app.get("/projects/<int:project_id>/export.xlsx")
@login_required
def export_project(project_id):
    p=project_or_404(project_id)
    with db() as conn:
        tasks=conn.execute("SELECT * FROM tasks WHERE project_id=? ORDER BY wbs,id",(project_id,)).fetchall()
        risks=conn.execute("SELECT * FROM risks WHERE project_id=? ORDER BY id",(project_id,)).fetchall()
        baselines=conn.execute("SELECT * FROM baselines WHERE project_id=? ORDER BY id",(project_id,)).fetchall()
        members=conn.execute("""SELECT pm.*,u.username,u.display_name,u.role global_role FROM project_members pm
                                JOIN users u ON u.id=pm.user_id WHERE pm.project_id=?""",(project_id,)).fetchall()
    wb=build_workbook(p,tasks,risks,baselines,members); bio=BytesIO(); wb.save(bio); bio.seek(0)
    return send_file(bio,as_attachment=True,download_name=f"{p['name']}-project-plan.xlsx",mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

@app.route("/admin")
@login_required
@role_required("admin")
def admin_center():
    with db() as conn:
        counts={
            "users":conn.execute("SELECT COUNT(*) c FROM users").fetchone()["c"],
            "active_users":conn.execute("SELECT COUNT(*) c FROM users WHERE active=1").fetchone()["c"],
            "projects":conn.execute("SELECT COUNT(*) c FROM projects").fetchone()["c"],
            "memberships":conn.execute("SELECT COUNT(*) c FROM project_members").fetchone()["c"],
        }
        recent=conn.execute("""SELECT a.*,u.display_name,p.name project_name FROM audit_log a
                               LEFT JOIN users u ON u.id=a.user_id LEFT JOIN projects p ON p.id=a.project_id
                               ORDER BY a.id DESC LIMIT 50""").fetchall()
    return render_template("admin_center.html",counts=counts,recent=recent)

@app.route("/admin/users",methods=["GET","POST"])
@login_required
@role_required("admin")
def admin_users():
    if request.method=="POST":
        username=request.form["username"].strip(); password=request.form["password"]; role=request.form.get("role","member")
        if role not in ROLES:role="member"
        if len(password)<10:
            flash("Lösenord måste vara minst 10 tecken.","danger"); return redirect(url_for("admin_users"))
        try:
            with db() as conn:
                conn.execute("""INSERT INTO users(username,display_name,password_hash,role,active,force_password_change,created_at)
                                VALUES(?,?,?,?,1,1,?)""",
                             (username,request.form.get("display_name","").strip() or username,
                              generate_password_hash(password),role,datetime.now().isoformat(timespec="seconds")))
            flash("Användaren skapades och måste byta lösenord vid första login.","success")
        except sqlite3.IntegrityError:
            flash("Användarnamnet finns redan.","danger")
        return redirect(url_for("admin_users"))
    with db() as conn: users=conn.execute("SELECT * FROM users ORDER BY username").fetchall()
    return render_template("users.html",users=users,roles=ROLES)

@app.post("/admin/users/<int:user_id>/toggle")
@login_required
@role_required("admin")
def toggle_user(user_id):
    u=current_user()
    if u["id"]==user_id:
        flash("Du kan inte inaktivera ditt eget konto.","danger")
    else:
        with db() as conn:
            row=conn.execute("SELECT active FROM users WHERE id=?",(user_id,)).fetchone()
            if row:conn.execute("UPDATE users SET active=? WHERE id=?",(0 if row["active"] else 1,user_id))
    return redirect(url_for("admin_users"))

@app.post("/admin/users/<int:user_id>/reset-password")
@login_required
@role_required("admin")
def admin_reset_password(user_id):
    new=request.form["password"]
    if len(new)<10:
        flash("Lösenord måste vara minst 10 tecken.","danger")
    else:
        with db() as conn:
            conn.execute("UPDATE users SET password_hash=?,force_password_change=1,failed_logins=0,locked_until='' WHERE id=?",
                         (generate_password_hash(new),user_id))
        flash("Lösenordet är återställt. Användaren måste byta det vid nästa login.","success")
    return redirect(url_for("admin_users"))

@app.post("/admin/users/<int:user_id>/role")
@login_required
@role_required("admin")
def admin_change_role(user_id):
    role=request.form.get("role","member")
    if role not in ROLES:abort(400)
    with db() as conn:conn.execute("UPDATE users SET role=? WHERE id=?",(role,user_id))
    return redirect(url_for("admin_users"))

@app.route("/admin/projects")
@login_required
@role_required("admin")
def admin_projects():
    with db() as conn:
        projects=conn.execute("""SELECT p.*,COUNT(pm.user_id) member_count FROM projects p
                                 LEFT JOIN project_members pm ON pm.project_id=p.id GROUP BY p.id ORDER BY p.id DESC""").fetchall()
    return render_template("admin_projects.html",projects=projects)

@app.post("/admin/projects/<int:project_id>/delete")
@login_required
@role_required("admin")
def admin_delete_project(project_id):
    with db() as conn:conn.execute("DELETE FROM projects WHERE id=?",(project_id,))
    flash("Projektet är borttaget.","success")
    return redirect(url_for("admin_projects"))

@app.route("/admin/settings",methods=["GET","POST"])
@login_required
@role_required("admin")
def admin_settings():
    if request.method=="POST":
        updates={
            "lockout_attempts":str(max(1,min(20,int(request.form.get("lockout_attempts","5"))))),
            "lockout_minutes":str(max(1,min(1440,int(request.form.get("lockout_minutes","15"))))),
        }
        with db() as conn:
            for k,v in updates.items():
                conn.execute("INSERT OR REPLACE INTO system_settings(key,value) VALUES(?,?)",(k,v))
        flash("Systeminställningar sparade.","success")
        return redirect(url_for("admin_settings"))
    return render_template("admin_settings.html",
                           lockout_attempts=setting("lockout_attempts","5"),
                           lockout_minutes=setting("lockout_minutes","15"))

@app.errorhandler(400)
def bad_request(e):return render_template("error.html",code=400,message=getattr(e,"description","Ogiltig begäran.")),400
@app.errorhandler(403)
def forbidden(e):return render_template("error.html",code=403,message="Du saknar behörighet för den här åtgärden."),403
@app.errorhandler(404)
def not_found(e):return render_template("error.html",code=404,message="Sidan kunde inte hittas."),404


def visible_projects_for_user():
    u=current_user()
    with db() as conn:
        if u["role"]=="admin":
            return conn.execute("SELECT * FROM projects ORDER BY name").fetchall()
        return conn.execute("""SELECT p.* FROM projects p JOIN project_members pm ON pm.project_id=p.id
                              WHERE pm.user_id=? ORDER BY p.name""",(u["id"],)).fetchall()

@app.get("/my-work")
@login_required
def my_work():
    u=current_user()
    with db() as conn:
        if u["role"]=="admin":
            rows=conn.execute("""SELECT t.*,p.name project_name FROM tasks t JOIN projects p ON p.id=t.project_id
                                 WHERE t.owner=? ORDER BY t.end_date,t.priority""",(u["display_name"],)).fetchall()
        else:
            rows=conn.execute("""SELECT t.*,p.name project_name FROM tasks t JOIN projects p ON p.id=t.project_id
                                 JOIN project_members pm ON pm.project_id=p.id AND pm.user_id=?
                                 WHERE t.owner IN (?,?) ORDER BY t.end_date,t.priority""",
                              (u["id"],u["display_name"],u["username"])).fetchall()
    items=[dict(r,derived_status=derived_status(r),health=task_health(r)) for r in rows]
    return render_template("my_work.html",items=items)

@app.get("/projects/<int:project_id>/kanban")
@login_required
def kanban(project_id):
    p=project_or_404(project_id)
    with db() as conn:
        tasks=conn.execute("SELECT * FROM tasks WHERE project_id=? ORDER BY sort_order,wbs,id",(project_id,)).fetchall()
    cols={s:[] for s in STATUSES}
    for t in tasks: cols.setdefault(t["status"],[]).append(t)
    return render_template("kanban.html",project=p,columns=cols,project_role=project_access(project_id))

@app.post("/projects/<int:project_id>/tasks/<int:task_id>/status")
@login_required
def set_task_status(project_id,task_id):
    project_or_404(project_id,write=True); task_or_404(project_id,task_id,write=True)
    status=request.form.get("status","Ej startad")
    if status not in STATUSES: abort(400)
    with db() as conn:
        conn.execute("UPDATE tasks SET status=?,progress=CASE WHEN ?='Klar' THEN 100 ELSE progress END WHERE id=?",
                     (status,status,task_id))
    audit(project_id,"task",task_id,"status",status)
    return redirect(request.referrer or url_for("kanban",project_id=project_id))

@app.get("/projects/<int:project_id>/calendar")
@login_required
def project_calendar(project_id):
    p=project_or_404(project_id)
    with db() as conn:
        tasks=conn.execute("""SELECT * FROM tasks WHERE project_id=? AND (start_date<>'' OR end_date<>'')
                              ORDER BY COALESCE(NULLIF(start_date,''),end_date),wbs""",(project_id,)).fetchall()
    return render_template("calendar.html",project=p,tasks=tasks)

def compute_cpm(project_id):
    with db() as conn:
        tasks=[dict(r) for r in conn.execute("SELECT * FROM tasks WHERE project_id=?",(project_id,))]
        links=[dict(r) for r in conn.execute("SELECT * FROM task_links WHERE project_id=? AND link_type='FS'",(project_id,))]
    if not tasks: return []
    ids={t["id"]:t for t in tasks}
    pred={i:[] for i in ids}; succ={i:[] for i in ids}
    for l in links:
        if l["predecessor_id"] in ids and l["successor_id"] in ids:
            pred[l["successor_id"]].append((l["predecessor_id"],int(l["lag_days"] or 0)))
            succ[l["predecessor_id"]].append((l["successor_id"],int(l["lag_days"] or 0)))
    indeg={i:len(pred[i]) for i in ids}; q=[i for i in ids if indeg[i]==0]; order=[]
    while q:
        i=q.pop(0); order.append(i)
        for j,_ in succ[i]:
            indeg[j]-=1
            if indeg[j]==0:q.append(j)
    if len(order)!=len(ids): return [{"error":"Cyclic dependency detected"}]
    dur={}
    for i,t in ids.items():
        s=parse_date(t["start_date"]); e=parse_date(t["end_date"])
        dur[i]=max(((e-s).days+1) if s and e else int(t.get("duration_days") or 1),1)
    es={i:0 for i in ids}; ef={}
    for i in order:
        es[i]=max([ef[p]+lag for p,lag in pred[i]] or [0]); ef[i]=es[i]+dur[i]
    finish=max(ef.values()); lf={i:finish for i in ids}; ls={}
    for i in reversed(order):
        if succ[i]: lf[i]=min(ls[j]-lag for j,lag in succ[i])
        ls[i]=lf[i]-dur[i]
    out=[]
    for i in order:
        t=ids[i]; slack=ls[i]-es[i]
        out.append(dict(t,early_start=es[i],early_finish=ef[i],late_start=ls[i],late_finish=lf[i],slack=slack,critical=(slack==0)))
    return out

@app.get("/projects/<int:project_id>/critical-path")
@login_required
def critical_path(project_id):
    p=project_or_404(project_id)
    return render_template("critical_path.html",project=p,rows=compute_cpm(project_id))

@app.post("/projects/<int:project_id>/links")
@login_required
def add_task_link(project_id):
    project_or_404(project_id,manager=True)
    pred=int(request.form["predecessor_id"]); succ=int(request.form["successor_id"])
    ltype=request.form.get("link_type","FS")
    if ltype not in ("FS","SS","FF","SF"): abort(400)
    if pred==succ: abort(400)
    with db() as conn:
        conn.execute("""INSERT INTO task_links(project_id,predecessor_id,successor_id,link_type,lag_days)
                        VALUES(?,?,?,?,?)""",(project_id,pred,succ,ltype,int(request.form.get("lag_days","0") or 0)))
    audit(project_id,"dependency",None,"created",f"{pred}->{succ} {ltype}")
    return redirect(url_for("critical_path",project_id=project_id))

@app.post("/projects/<int:project_id>/auto-schedule")
@login_required
def auto_schedule(project_id):
    project_or_404(project_id,manager=True)
    with db() as conn:
        tasks={r["id"]:dict(r) for r in conn.execute("SELECT * FROM tasks WHERE project_id=?",(project_id,))}
        links=[dict(r) for r in conn.execute("SELECT * FROM task_links WHERE project_id=? AND link_type='FS'",(project_id,))]
        changed=0
        for _ in range(max(len(tasks),1)):
            round_changed=0
            for l in links:
                a=tasks.get(l["predecessor_id"]); b=tasks.get(l["successor_id"])
                if not a or not b: continue
                ae=parse_date(a["end_date"]); bs=parse_date(b["start_date"]); be=parse_date(b["end_date"])
                if not ae or not bs: continue
                target=ae+timedelta(days=1+int(l["lag_days"] or 0))
                if bs<target:
                    dur=max(((be-bs).days if be else 0),0)
                    newe=target+timedelta(days=dur)
                    conn.execute("UPDATE tasks SET start_date=?,end_date=? WHERE id=?",(target.isoformat(),newe.isoformat(),b["id"]))
                    b["start_date"]=target.isoformat(); b["end_date"]=newe.isoformat()
                    round_changed+=1; changed+=1
            if not round_changed: break
    audit(project_id,"schedule",None,"auto-schedule",f"{changed} flyttningar")
    flash(f"Auto-scheduling klar: {changed} datumjusteringar.","success")
    return redirect(url_for("gantt",project_id=project_id))


@app.get("/resources")
@login_required
def resources():
    u=current_user()
    with db() as conn:
        if u["role"]=="admin":
            users=conn.execute("""SELECT u.*,COALESCE(r.weekly_capacity,40) weekly_capacity,COALESCE(r.hourly_cost,0) hourly_cost
                                  FROM users u LEFT JOIN resource_profiles r ON r.user_id=u.id WHERE u.active=1 ORDER BY u.display_name""").fetchall()
        else:
            users=conn.execute("""SELECT u.*,COALESCE(r.weekly_capacity,40) weekly_capacity,COALESCE(r.hourly_cost,0) hourly_cost
                                  FROM users u LEFT JOIN resource_profiles r ON r.user_id=u.id
                                  WHERE u.id IN (SELECT DISTINCT pm2.user_id FROM project_members pm1
                                    JOIN project_members pm2 ON pm2.project_id=pm1.project_id WHERE pm1.user_id=?)
                                  ORDER BY u.display_name""",(u["id"],)).fetchall()
        loads=[]
        for person in users:
            hrs=conn.execute("""SELECT COALESCE(SUM(t.planned_hours),0) h FROM tasks t
                                WHERE t.owner IN (?,?) AND t.status<>'Klar'""",(person["display_name"],person["username"])).fetchone()["h"]
            loads.append(dict(person,planned_open=float(hrs or 0),utilization=round(float(hrs or 0)/max(float(person["weekly_capacity"] or 40),1)*100)))
    return render_template("resources.html",resources=loads)

@app.post("/admin/resources/<int:user_id>")
@login_required
@role_required("admin")
def update_resource(user_id):
    with db() as conn:
        conn.execute("""INSERT INTO resource_profiles(user_id,weekly_capacity,hourly_cost,hourly_rate)
                        VALUES(?,?,?,?) ON CONFLICT(user_id) DO UPDATE SET weekly_capacity=excluded.weekly_capacity,
                        hourly_cost=excluded.hourly_cost,hourly_rate=excluded.hourly_rate""",
                     (user_id,float(request.form.get("weekly_capacity","40")),float(request.form.get("hourly_cost","0")),
                      float(request.form.get("hourly_rate","0"))))
    return redirect(url_for("resources"))

@app.route("/projects/<int:project_id>/time",methods=["GET","POST"])
@login_required
def project_time(project_id):
    p=project_or_404(project_id,write=(request.method=="POST"))
    u=current_user()
    with db() as conn:
        tasks=conn.execute("SELECT id,wbs,title FROM tasks WHERE project_id=? ORDER BY wbs",(project_id,)).fetchall()
        if request.method=="POST":
            conn.execute("""INSERT INTO time_entries(project_id,task_id,user_id,work_date,hours,note,created_at)
                            VALUES(?,?,?,?,?,?,?)""",(project_id,request.form.get("task_id") or None,u["id"],
                            request.form.get("work_date") or date.today().isoformat(),float(request.form["hours"]),
                            request.form.get("note","").strip(),datetime.now().isoformat(timespec="seconds")))
            audit(project_id,"time",None,"logged",f"{request.form['hours']} h")
            flash("Tid rapporterad.","success")
            return redirect(url_for("project_time",project_id=project_id))
        entries=conn.execute("""SELECT te.*,u.display_name,t.title task_title FROM time_entries te
                                JOIN users u ON u.id=te.user_id LEFT JOIN tasks t ON t.id=te.task_id
                                WHERE te.project_id=? ORDER BY te.work_date DESC,te.id DESC""",(project_id,)).fetchall()
        total=conn.execute("SELECT COALESCE(SUM(hours),0) h FROM time_entries WHERE project_id=?",(project_id,)).fetchone()["h"]
    return render_template("time.html",project=p,tasks=tasks,entries=entries,total=total)

@app.route("/projects/<int:project_id>/financials",methods=["GET","POST"])
@login_required
def financials(project_id):
    p=project_or_404(project_id,manager=(request.method=="POST"))
    with db() as conn:
        if request.method=="POST":
            if request.form.get("action")=="budget":
                conn.execute("""INSERT INTO project_finance(project_id,budget,currency) VALUES(?,?,?)
                                ON CONFLICT(project_id) DO UPDATE SET budget=excluded.budget,currency=excluded.currency""",
                             (project_id,float(request.form.get("budget","0")),request.form.get("currency","SEK")))
            else:
                conn.execute("""INSERT INTO project_costs(project_id,category,description,planned,actual,cost_date)
                                VALUES(?,?,?,?,?,?)""",(project_id,request.form.get("category","External"),request.form["description"],
                                float(request.form.get("planned","0")),float(request.form.get("actual","0")),request.form.get("cost_date","")))
            audit(project_id,"finance",None,"updated","")
            return redirect(url_for("financials",project_id=project_id))
        finance=conn.execute("SELECT * FROM project_finance WHERE project_id=?",(project_id,)).fetchone()
        costs=conn.execute("SELECT * FROM project_costs WHERE project_id=? ORDER BY id DESC",(project_id,)).fetchall()
        time_cost=conn.execute("""SELECT COALESCE(SUM(te.hours*COALESCE(r.hourly_cost,0)),0) c FROM time_entries te
                                  LEFT JOIN resource_profiles r ON r.user_id=te.user_id WHERE te.project_id=?""",(project_id,)).fetchone()["c"]
    budget=float(finance["budget"] if finance else 0); external=sum(float(c["actual"] or 0) for c in costs)
    actual=external+float(time_cost or 0)
    return render_template("financials.html",project=p,finance=finance,costs=costs,time_cost=time_cost,actual=actual,
                           variance=budget-actual,budget=budget)

@app.get("/projects/<int:project_id>/earned-value")
@login_required
def earned_value(project_id):
    p=project_or_404(project_id)
    with db() as conn:
        tasks=conn.execute("SELECT * FROM tasks WHERE project_id=?",(project_id,)).fetchall()
        fin=conn.execute("SELECT * FROM project_finance WHERE project_id=?",(project_id,)).fetchone()
        actual_external=conn.execute("SELECT COALESCE(SUM(actual),0) a FROM project_costs WHERE project_id=?",(project_id,)).fetchone()["a"]
        time_cost=conn.execute("""SELECT COALESCE(SUM(te.hours*COALESCE(r.hourly_cost,0)),0) c FROM time_entries te
                                  LEFT JOIN resource_profiles r ON r.user_id=te.user_id WHERE te.project_id=?""",(project_id,)).fetchone()["c"]
    bac=float(fin["budget"] if fin else 0)
    today=date.today(); weighted_plan=0; weighted_ev=0
    for t in tasks:
        s=parse_date(t["start_date"]); e=parse_date(t["end_date"])
        if s and e:
            if today>=e: planned=1
            elif today<=s: planned=0
            else: planned=(today-s).days/max((e-s).days,1)
        else: planned=0
        weighted_plan+=planned; weighted_ev+=float(t["progress"] or 0)/100
    n=max(len(tasks),1); pv=bac*(weighted_plan/n); ev=bac*(weighted_ev/n); ac=float(actual_external or 0)+float(time_cost or 0)
    spi=(ev/pv) if pv else 0; cpi=(ev/ac) if ac else 0
    return render_template("earned_value.html",project=p,bac=bac,pv=pv,ev=ev,ac=ac,spi=spi,cpi=cpi)


@app.route("/pmo",methods=["GET","POST"])
@login_required
def pmo():
    u=current_user()
    with db() as conn:
        if request.method=="POST":
            if u["role"] not in ("admin","pm"): abort(403)
            conn.execute("INSERT INTO portfolios(name,description,owner) VALUES(?,?,?)",
                         (request.form["name"],request.form.get("description",""),u["display_name"]))
            return redirect(url_for("pmo"))
        portfolios=conn.execute("""SELECT pf.*,COUNT(pp.project_id) project_count FROM portfolios pf
                                   LEFT JOIN portfolio_projects pp ON pp.portfolio_id=pf.id GROUP BY pf.id ORDER BY pf.name""").fetchall()
        latest=conn.execute("""SELECT sr.*,p.name project_name FROM status_reports sr JOIN projects p ON p.id=sr.project_id
                               WHERE sr.id IN (SELECT MAX(id) FROM status_reports GROUP BY project_id)
                               ORDER BY sr.report_date DESC""").fetchall()
    return render_template("pmo.html",portfolios=portfolios,latest=latest)

@app.route("/projects/<int:project_id>/status-report",methods=["GET","POST"])
@login_required
def status_report(project_id):
    p=project_or_404(project_id,manager=(request.method=="POST"))
    u=current_user()
    with db() as conn:
        if request.method=="POST":
            conn.execute("""INSERT INTO status_reports(project_id,report_date,overall_rag,scope_rag,schedule_rag,budget_rag,resources_rag,
                            summary,achievements,next_steps,created_by,created_at) VALUES(?,?,?,?,?,?,?,?,?,?,?,?)""",
                         (project_id,date.today().isoformat(),request.form["overall_rag"],request.form["scope_rag"],
                          request.form["schedule_rag"],request.form["budget_rag"],request.form["resources_rag"],
                          request.form.get("summary",""),request.form.get("achievements",""),request.form.get("next_steps",""),
                          u["id"],datetime.now().isoformat(timespec="seconds")))
            return redirect(url_for("status_report",project_id=project_id))
        reports=conn.execute("SELECT * FROM status_reports WHERE project_id=? ORDER BY id DESC",(project_id,)).fetchall()
    return render_template("status_report.html",project=p,reports=reports)

@app.route("/projects/<int:project_id>/governance",methods=["GET","POST"])
@login_required
def governance(project_id):
    p=project_or_404(project_id,write=(request.method=="POST"))
    u=current_user()
    with db() as conn:
        if request.method=="POST":
            kind=request.form["kind"]
            if kind=="raid":
                conn.execute("""INSERT INTO raid_items(project_id,item_type,title,owner,status,due_date,details) VALUES(?,?,?,?,?,?,?)""",
                             (project_id,request.form["item_type"],request.form["title"],request.form.get("owner",""),
                              request.form.get("status","Open"),request.form.get("due_date",""),request.form.get("details","")))
            elif kind=="decision":
                conn.execute("""INSERT INTO decisions(project_id,title,decision,decided_by,decision_date) VALUES(?,?,?,?,?)""",
                             (project_id,request.form["title"],request.form["decision"],request.form.get("decided_by",u["display_name"]),
                              request.form.get("decision_date") or date.today().isoformat()))
            elif kind=="change":
                conn.execute("""INSERT INTO change_requests(project_id,title,description,impact_scope,impact_days,impact_cost,status,requested_by,created_at)
                                VALUES(?,?,?,?,?,?,?,?,?)""",(project_id,request.form["title"],request.form.get("description",""),
                                request.form.get("impact_scope",""),int(request.form.get("impact_days","0") or 0),
                                float(request.form.get("impact_cost","0") or 0),"Proposed",u["id"],datetime.now().isoformat(timespec="seconds")))
            audit(project_id,kind,None,"created",request.form.get("title",""))
            return redirect(url_for("governance",project_id=project_id))
        raid=conn.execute("SELECT * FROM raid_items WHERE project_id=? ORDER BY id DESC",(project_id,)).fetchall()
        decisions=conn.execute("SELECT * FROM decisions WHERE project_id=? ORDER BY decision_date DESC,id DESC",(project_id,)).fetchall()
        changes=conn.execute("SELECT * FROM change_requests WHERE project_id=? ORDER BY id DESC",(project_id,)).fetchall()
        approvals=conn.execute("SELECT * FROM approvals WHERE project_id=? ORDER BY id DESC",(project_id,)).fetchall()
    return render_template("governance.html",project=p,raid=raid,decisions=decisions,changes=changes,approvals=approvals)

@app.post("/projects/<int:project_id>/changes/<int:change_id>/decision")
@login_required
def decide_change(project_id,change_id):
    project_or_404(project_id,manager=True)
    status=request.form.get("status")
    if status not in ("Approved","Rejected"): abort(400)
    with db() as conn: conn.execute("UPDATE change_requests SET status=? WHERE id=? AND project_id=?",(status,change_id,project_id))
    audit(project_id,"change",change_id,status,"")
    return redirect(url_for("governance",project_id=project_id))


UPLOAD_DIR=DATA_DIR/"uploads"

@app.get("/notifications")
@login_required
def notifications_page():
    u=current_user()
    with db() as conn:
        items=conn.execute("SELECT * FROM notifications WHERE user_id=? ORDER BY id DESC LIMIT 100",(u["id"],)).fetchall()
        conn.execute("UPDATE notifications SET is_read=1 WHERE user_id=?",(u["id"],))
    return render_template("notifications.html",items=items)

@app.route("/projects/<int:project_id>/collaboration",methods=["GET","POST"])
@login_required
def collaboration(project_id):
    p=project_or_404(project_id,write=(request.method=="POST"))
    u=current_user()
    with db() as conn:
        if request.method=="POST":
            kind=request.form["kind"]
            if kind=="doc":
                conn.execute("INSERT INTO documents(project_id,title,body,updated_by,updated_at) VALUES(?,?,?,?,?)",
                             (project_id,request.form["title"],request.form.get("body",""),u["id"],datetime.now().isoformat(timespec="seconds")))
            elif kind=="meeting":
                conn.execute("INSERT INTO meetings(project_id,title,meeting_date,attendees,notes) VALUES(?,?,?,?,?)",
                             (project_id,request.form["title"],request.form.get("meeting_date") or date.today().isoformat(),
                              request.form.get("attendees",""),request.form.get("notes","")))
            elif kind=="action":
                conn.execute("INSERT INTO action_items(project_id,title,owner,due_date,status) VALUES(?,?,?,?,?)",
                             (project_id,request.form["title"],request.form.get("owner",""),request.form.get("due_date",""),"Open"))
            audit(project_id,kind,None,"created",request.form.get("title",""))
            return redirect(url_for("collaboration",project_id=project_id))
        docs=conn.execute("SELECT * FROM documents WHERE project_id=? ORDER BY updated_at DESC",(project_id,)).fetchall()
        meetings=conn.execute("SELECT * FROM meetings WHERE project_id=? ORDER BY meeting_date DESC,id DESC",(project_id,)).fetchall()
        actions=conn.execute("SELECT * FROM action_items WHERE project_id=? ORDER BY status,due_date",(project_id,)).fetchall()
        files=conn.execute("SELECT * FROM attachments WHERE project_id=? ORDER BY id DESC",(project_id,)).fetchall()
    return render_template("collaboration.html",project=p,docs=docs,meetings=meetings,actions=actions,files=files)

@app.post("/projects/<int:project_id>/upload")
@login_required
def upload_attachment(project_id):
    project_or_404(project_id,write=True)
    f=request.files.get("file")
    if not f or not f.filename: abort(400)
    UPLOAD_DIR.mkdir(parents=True,exist_ok=True)
    original=secure_filename(f.filename)
    stored=f"{project_id}-{secrets.token_hex(8)}-{original}"
    f.save(UPLOAD_DIR/stored)
    u=current_user()
    with db() as conn:
        conn.execute("""INSERT INTO attachments(project_id,entity_type,original_name,stored_name,uploaded_by,uploaded_at)
                        VALUES(?,?,?,?,?,?)""",(project_id,"project",original,stored,u["id"],datetime.now().isoformat(timespec="seconds")))
    audit(project_id,"attachment",None,"uploaded",original)
    return redirect(url_for("collaboration",project_id=project_id))

@app.get("/projects/<int:project_id>/files/<int:file_id>")
@login_required
def download_attachment(project_id,file_id):
    project_or_404(project_id)
    with db() as conn: row=conn.execute("SELECT * FROM attachments WHERE id=? AND project_id=?",(file_id,project_id)).fetchone()
    if not row: abort(404)
    return send_file(UPLOAD_DIR/row["stored_name"],as_attachment=True,download_name=row["original_name"])

@app.post("/projects/<int:project_id>/notify")
@login_required
def notify_project(project_id):
    project_or_404(project_id,manager=True)
    msg=request.form["message"].strip()
    with db() as conn:
        members=conn.execute("SELECT user_id FROM project_members WHERE project_id=?",(project_id,)).fetchall()
        for m in members:
            conn.execute("INSERT INTO notifications(user_id,project_id,message,link,is_read,created_at) VALUES(?,?,?,?,0,?)",
                         (m["user_id"],project_id,msg,url_for("project",project_id=project_id),datetime.now().isoformat(timespec="seconds")))
    return redirect(url_for("collaboration",project_id=project_id))


BACKUP_DIR=DATA_DIR/"admin-backups"

def api_user():
    auth=request.headers.get("Authorization","")
    if not auth.startswith("Bearer "): return None
    raw=auth[7:].strip(); digest=hashlib.sha256(raw.encode()).hexdigest()
    with db() as conn:
        row=conn.execute("""SELECT u.*,k.id key_id FROM api_keys k JOIN users u ON u.id=k.user_id
                            WHERE k.key_hash=? AND k.active=1 AND u.active=1""",(digest,)).fetchone()
        if row: conn.execute("UPDATE api_keys SET last_used=? WHERE id=?",(datetime.now().isoformat(timespec="seconds"),row["key_id"]))
    return row

def api_required(fn):
    @wraps(fn)
    def wrap(*args,**kwargs):
        u=api_user()
        if not u:return jsonify(error="unauthorized"),401
        request.api_user=u
        return fn(*args,**kwargs)
    return wrap

@app.get("/api/v1/projects")
@api_required
def api_projects():
    u=request.api_user
    with db() as conn:
        if u["role"]=="admin":
            rows=conn.execute("SELECT * FROM projects ORDER BY id").fetchall()
        else:
            rows=conn.execute("""SELECT p.* FROM projects p JOIN project_members pm ON pm.project_id=p.id
                                 WHERE pm.user_id=? ORDER BY p.id""",(u["id"],)).fetchall()
    return jsonify([dict(r) for r in rows])

@app.get("/api/v1/projects/<int:project_id>/tasks")
@api_required
def api_tasks(project_id):
    u=request.api_user
    with db() as conn:
        if u["role"]!="admin":
            ok=conn.execute("SELECT 1 FROM project_members WHERE project_id=? AND user_id=?",(project_id,u["id"])).fetchone()
            if not ok:return jsonify(error="forbidden"),403
        rows=conn.execute("SELECT * FROM tasks WHERE project_id=? ORDER BY wbs,id",(project_id,)).fetchall()
    return jsonify([dict(r) for r in rows])

@app.route("/admin/api-keys",methods=["GET","POST"])
@login_required
@role_required("admin")
def admin_api_keys():
    generated=None
    with db() as conn:
        if request.method=="POST":
            raw="pp_"+secrets.token_urlsafe(30); digest=hashlib.sha256(raw.encode()).hexdigest()
            conn.execute("INSERT INTO api_keys(user_id,name,key_hash,prefix,created_at) VALUES(?,?,?,?,?)",
                         (int(request.form["user_id"]),request.form["name"],digest,raw[:10],datetime.now().isoformat(timespec="seconds")))
            generated=raw
        keys=conn.execute("""SELECT k.*,u.username FROM api_keys k JOIN users u ON u.id=k.user_id ORDER BY k.id DESC""").fetchall()
        users=conn.execute("SELECT * FROM users WHERE active=1 ORDER BY username").fetchall()
    return render_template("api_keys.html",keys=keys,users=users,generated=generated)

@app.route("/admin/custom-fields",methods=["GET","POST"])
@login_required
@role_required("admin")
def admin_custom_fields():
    with db() as conn:
        if request.method=="POST":
            conn.execute("""INSERT INTO custom_field_definitions(entity_type,name,field_type,options_json,required)
                            VALUES(?,?,?,?,?)""",(request.form["entity_type"],request.form["name"],request.form["field_type"],
                            json.dumps([x.strip() for x in request.form.get("options","").split(",") if x.strip()]),
                            1 if request.form.get("required")=="on" else 0))
            return redirect(url_for("admin_custom_fields"))
        fields=conn.execute("SELECT * FROM custom_field_definitions ORDER BY entity_type,name").fetchall()
    return render_template("custom_fields.html",fields=fields)

@app.route("/admin/backups",methods=["GET","POST"])
@login_required
@role_required("admin")
def admin_backups():
    BACKUP_DIR.mkdir(parents=True,exist_ok=True)
    if request.method=="POST":
        name=f"projectplan-{datetime.now():%Y%m%d-%H%M%S}.db"
        shutil.copy2(DB_PATH,BACKUP_DIR/name)
        flash("Backup skapad.","success")
        return redirect(url_for("admin_backups"))
    files=sorted(BACKUP_DIR.glob("*.db"),reverse=True)
    return render_template("backups.html",files=[{"name":f.name,"size":f.stat().st_size} for f in files])

@app.get("/admin/backups/<name>")
@login_required
@role_required("admin")
def download_backup(name):
    safe=Path(name).name
    return send_file(BACKUP_DIR/safe,as_attachment=True,download_name=safe)

@app.get("/admin/diagnostics")
@login_required
@role_required("admin")
def diagnostics():
    with db() as conn:
        db_size=DB_PATH.stat().st_size if DB_PATH.exists() else 0
        tables=conn.execute("SELECT name FROM sqlite_master WHERE type='table' ORDER BY name").fetchall()
        migrations=conn.execute("SELECT * FROM schema_migrations ORDER BY applied_at DESC").fetchall()
    info={"version":APP_VERSION,"python":platform.python_version(),"platform":platform.platform(),
          "db_path":str(DB_PATH),"db_size":db_size,"table_count":len(tables),"upload_dir":str(UPLOAD_DIR)}
    return render_template("diagnostics.html",info=info,migrations=migrations)

@app.get("/api")
@login_required
def api_docs():
    return render_template("api_docs.html")


@app.get("/about")
@login_required
def about():
    return render_template("about.html")

def roadmap_accessible_projects():
    u=current_user()
    with db() as conn:
        if u["role"]=="admin":
            return conn.execute("SELECT * FROM projects ORDER BY name").fetchall()
        return conn.execute("""SELECT p.* FROM projects p JOIN project_members pm ON pm.project_id=p.id
                              WHERE pm.user_id=? ORDER BY p.name""",(u["id"],)).fetchall()

def roadmap_task_tree(tasks):
    rows=[dict(t) for t in tasks]
    by_id={int(t["id"]):t for t in rows}
    children={}
    roots=[]
    for t in rows:
        pid=t.get("parent_task_id")
        if pid and int(pid) in by_id and int(pid)!=int(t["id"]):
            children.setdefault(int(pid),[]).append(t)
        else:
            roots.append(t)
    def key(t):
        raw=str(t.get("wbs") or "")
        parts=[]
        for p in raw.split("."):
            try: parts.append((0,int(p)))
            except: parts.append((1,p.lower()))
        return (parts,int(t.get("sort_order") or 0),int(t["id"]))
    for vals in children.values(): vals.sort(key=key)
    roots.sort(key=key)
    out=[]
    def walk(t,depth):
        r=dict(t); r["depth"]=depth; r["derived"]=derived_status(t); r["children_count"]=len(children.get(int(t["id"]),[]))
        out.append(r)
        for c in children.get(int(t["id"]),[]): walk(c,depth+1)
    for r in roots: walk(r,0)
    return out

@app.get("/planning")
@login_required
def planning_hub():
    projects=roadmap_accessible_projects()
    return render_template("planning_hub.html",projects=projects)

@app.get("/projects/<int:project_id>/cockpit")
@login_required
def project_cockpit(project_id):
    p=project_or_404(project_id)
    today=date.today()
    with db() as conn:
        tasks=[dict(r) for r in conn.execute("SELECT * FROM tasks WHERE project_id=? ORDER BY wbs,id",(project_id,))]
        risks=[dict(r) for r in conn.execute("SELECT * FROM risks WHERE project_id=? AND status<>'Stängd' ORDER BY probability*impact DESC",(project_id,))]
        actions=[dict(r) for r in conn.execute("SELECT * FROM action_items WHERE project_id=? AND status<>'Done' ORDER BY due_date,id",(project_id,))]
        recent=conn.execute("""SELECT a.*,u.display_name FROM audit_log a LEFT JOIN users u ON u.id=a.user_id
                              WHERE a.project_id=? ORDER BY a.id DESC LIMIT 10""",(project_id,)).fetchall()
    overdue=[t for t in tasks if int(t.get("progress") or 0)<100 and parse_date(t.get("end_date")) and parse_date(t.get("end_date"))<today]
    blocked=[t for t in tasks if t.get("status")=="Blockerad" and int(t.get("progress") or 0)<100]
    high_risks=[r for r in risks if int(r.get("probability") or 0)*int(r.get("impact") or 0)>=12]
    milestones=[t for t in tasks if t.get("milestone") and parse_date(t.get("end_date")) and int(t.get("progress") or 0)<100]
    milestones.sort(key=lambda x:parse_date(x.get("end_date")))
    avg=round(sum(int(t.get("progress") or 0) for t in tasks)/len(tasks)) if tasks else 0
    rag="red" if overdue or len(high_risks)>=2 or len(blocked)>=2 else ("amber" if high_risks or blocked else "green")
    attention=[]
    for t in overdue[:5]: attention.append(("Försenad",t["title"],t.get("end_date") or "","danger"))
    for t in blocked[:5]: attention.append(("Blockerad",t["title"],t.get("owner") or "","warning"))
    for r in high_risks[:5]: attention.append(("Hög risk",r["title"],f"Score {int(r.get('probability') or 0)*int(r.get('impact') or 0)}","danger"))
    return render_template("cockpit.html",project=p,tasks=tasks,risks=risks,actions=actions,recent=recent,
                           overdue=overdue,blocked=blocked,high_risks=high_risks,milestones=milestones[:8],
                           avg=avg,rag=rag,attention=attention)

@app.get("/projects/<int:project_id>/wbs")
@login_required
def project_wbs(project_id):
    p=project_or_404(project_id)
    with db() as conn:
        tasks=conn.execute("SELECT * FROM tasks WHERE project_id=? ORDER BY wbs,sort_order,id",(project_id,)).fetchall()
    return render_template("wbs.html",project=p,rows=roadmap_task_tree(tasks))

@app.post("/projects/<int:project_id>/tasks/<int:task_id>/quick")
@login_required
def task_quick_update(project_id,task_id):
    project_or_404(project_id,write=True)
    progress=max(0,min(100,int(request.form.get("progress","0") or 0)))
    status=request.form.get("status","Pågår")
    if progress==100: status="Klar"
    with db() as conn:
        conn.execute("UPDATE tasks SET progress=?,status=? WHERE id=? AND project_id=?",(progress,status,task_id,project_id))
    audit(project_id,"task",task_id,"quick_updated",f"{status} {progress}%")
    return redirect(request.referrer or url_for("project_wbs",project_id=project_id))

@app.get("/my-work-2")
@login_required
def my_work_2():
    u=current_user(); today=date.today()
    with db() as conn:
        projects=roadmap_accessible_projects()
        ids=[p["id"] for p in projects]
        if not ids: tasks=[]; actions=[]
        else:
            ph=",".join("?" for _ in ids)
            tasks=[dict(r) for r in conn.execute(f"""SELECT t.*,p.name project_name FROM tasks t JOIN projects p ON p.id=t.project_id
                WHERE t.project_id IN ({ph}) AND t.progress<100 AND
                (LOWER(t.owner)=LOWER(?) OR t.owner_user_id=?) ORDER BY t.end_date,t.priority""",(*ids,u["display_name"],u["id"]))]
            actions=[dict(r) for r in conn.execute(f"""SELECT a.*,p.name project_name FROM action_items a JOIN projects p ON p.id=a.project_id
                WHERE a.project_id IN ({ph}) AND a.status<>'Done' AND LOWER(a.owner)=LOWER(?) ORDER BY a.due_date""",(*ids,u["display_name"]))]
    for t in tasks:
        d=parse_date(t.get("end_date")); t["is_overdue"]=bool(d and d<today); t["is_week"]=bool(d and today<=d<=today+timedelta(days=7))
    return render_template("my_work_2.html",tasks=tasks,actions=actions,today=today)

@app.get("/projects/<int:project_id>/gantt-pro")
@login_required
def gantt_pro(project_id):
    p=project_or_404(project_id)
    with db() as conn:
        tasks=[dict(r) for r in conn.execute("SELECT * FROM tasks WHERE project_id=? ORDER BY wbs,id",(project_id,))]
        links=[dict(r) for r in conn.execute("SELECT * FROM task_links WHERE project_id=? ORDER BY id",(project_id,))]
    dated=[t for t in tasks if parse_date(t.get("start_date")) and parse_date(t.get("end_date"))]
    if dated:
        rs=min(parse_date(t["start_date"]) for t in dated); re=max(parse_date(t["end_date"]) for t in dated)
        total=max(1,(re-rs).days+1)
        for t in dated:
            st=parse_date(t["start_date"]); en=parse_date(t["end_date"])
            t["left"]=round((st-rs).days/total*100,3); t["width"]=max(.8,round(((en-st).days+1)/total*100,3))
    else: rs=re=None
    return render_template("gantt_pro.html",project=p,tasks=roadmap_task_tree(tasks),dated=dated,links=links,rs=rs,re=re)

@app.route("/control-center",methods=["GET","POST"])
@login_required
def control_center():
    projects=roadmap_accessible_projects()
    project_id=request.args.get("project_id",type=int) or (projects[0]["id"] if projects else None)
    if project_id: project_or_404(project_id)
    with db() as conn:
        raid=conn.execute("SELECT * FROM raid_items WHERE project_id=? ORDER BY kind,status,due_date",(project_id,)).fetchall() if project_id else []
        changes=conn.execute("SELECT * FROM change_requests WHERE project_id=? ORDER BY id DESC",(project_id,)).fetchall() if project_id else []
        decisions=conn.execute("SELECT * FROM decisions WHERE project_id=? ORDER BY decided_at DESC,id DESC",(project_id,)).fetchall() if project_id else []
        approvals=conn.execute("SELECT * FROM approvals WHERE project_id=? ORDER BY id DESC",(project_id,)).fetchall() if project_id else []
    return render_template("control_center.html",projects=projects,project_id=project_id,raid=raid,changes=changes,decisions=decisions,approvals=approvals)

@app.post("/projects/<int:project_id>/raid/new")
@login_required
def roadmap_raid_new(project_id):
    project_or_404(project_id,write=True)
    with db() as conn:
        conn.execute("""INSERT INTO raid_items(project_id,item_type,title,owner,status,due_date,details)
                        VALUES(?,?,?,?,?,?,?)""",(project_id,request.form.get("kind","Risk"),request.form["title"].strip(),
                        request.form.get("owner","").strip(),request.form.get("status","Open"),request.form.get("due_date",""),
                        request.form.get("description","").strip()))
    audit(project_id,"raid",None,"created",request.form["title"].strip())
    return redirect(url_for("control_center",project_id=project_id))

@app.post("/projects/<int:project_id>/changes/new")
@login_required
def roadmap_change_new(project_id):
    project_or_404(project_id,write=True)
    u=current_user()
    with db() as conn:
        conn.execute("""INSERT INTO change_requests(project_id,title,description,reason,impact_cost,impact_days,status,requested_by,created_at)
                        VALUES(?,?,?,?,?,?,?,?,?)""",(project_id,request.form["title"].strip(),request.form.get("description","").strip(),
                        request.form.get("reason","").strip(),float(request.form.get("impact_cost","0") or 0),int(request.form.get("impact_days","0") or 0),
                        "Submitted",u["id"],datetime.now().isoformat(timespec="seconds")))
    audit(project_id,"change",None,"submitted",request.form["title"].strip())
    return redirect(url_for("control_center",project_id=project_id))

@app.post("/projects/<int:project_id>/changes/<int:change_id>/<decision>")
@login_required
def roadmap_change_decision(project_id,change_id,decision):
    project_or_404(project_id,manager=True)
    status="Approved" if decision=="approve" else "Rejected"
    with db() as conn:
        conn.execute("UPDATE change_requests SET status=?,decided_by=?,decided_at=? WHERE id=? AND project_id=?",
                     (status,current_user()["id"],datetime.now().isoformat(timespec="seconds"),change_id,project_id))
    audit(project_id,"change",change_id,status.lower(),"")
    return redirect(url_for("control_center",project_id=project_id))

@app.post("/projects/<int:project_id>/decisions/new")
@login_required
def roadmap_decision_new(project_id):
    project_or_404(project_id,write=True)
    with db() as conn:
        conn.execute("""INSERT INTO decisions(project_id,title,decision,decided_by,decision_date,owner,decided_at)
                        VALUES(?,?,?,?,?,?,?)""",(project_id,request.form["title"].strip(),request.form.get("decision","").strip(),
                        request.form.get("owner","").strip(),date.today().isoformat(),request.form.get("owner","").strip(),
                        datetime.now().isoformat(timespec="seconds")))
    return redirect(url_for("control_center",project_id=project_id))

@app.route("/resource-plan",methods=["GET","POST"])
@login_required
def resource_plan():
    projects=roadmap_accessible_projects()
    if request.method=="POST":
        pid=int(request.form["project_id"]); project_or_404(pid,manager=True)
        with db() as conn:
            conn.execute("""INSERT INTO resource_allocations(project_id,user_id,resource_name,week_start,allocation_pct,planned_hours)
                            VALUES(?,?,?,?,?,?)""",(pid,int(request.form["user_id"]) if request.form.get("user_id") else None,
                            request.form["resource_name"].strip(),request.form["week_start"],int(request.form.get("allocation_pct","0") or 0),
                            float(request.form.get("planned_hours","0") or 0)))
        return redirect(url_for("resource_plan"))
    ids=[p["id"] for p in projects]
    with db() as conn:
        users=conn.execute("SELECT * FROM users WHERE active=1 ORDER BY display_name").fetchall()
        if ids:
            ph=",".join("?" for _ in ids)
            allocations=conn.execute(f"""SELECT a.*,p.name project_name FROM resource_allocations a JOIN projects p ON p.id=a.project_id
                                        WHERE a.project_id IN ({ph}) ORDER BY a.week_start,a.resource_name""",ids).fetchall()
        else: allocations=[]
    totals={}
    for a in allocations:
        k=(a["resource_name"],a["week_start"])
        totals[k]=totals.get(k,0)+int(a["allocation_pct"] or 0)
    return render_template("resource_plan.html",projects=projects,users=users,allocations=allocations,totals=totals)

@app.get("/finance-control")
@login_required
def finance_control():
    projects=roadmap_accessible_projects(); ids=[p["id"] for p in projects]
    rows=[]
    with db() as conn:
        for p in projects:
            f=conn.execute("SELECT * FROM project_finance WHERE project_id=?",(p["id"],)).fetchone()
            ac=conn.execute("SELECT COALESCE(SUM(actual),0) v FROM project_costs WHERE project_id=?",(p["id"],)).fetchone()["v"]
            planned=conn.execute("SELECT COALESCE(SUM(planned),0) v FROM project_costs WHERE project_id=?",(p["id"],)).fetchone()["v"]
            budget=float(f["budget"] or 0) if f else 0
            actual=float(ac or 0); forecast=max(float(planned or 0),actual)
            rows.append({"project":p,"budget":budget,"actual":actual,"forecast":forecast,
                         "variance":budget-forecast,"used_pct":round(actual/budget*100) if budget else 0})
    return render_template("finance_control.html",rows=rows)

@app.route("/executive-pmo",methods=["GET","POST"])
@login_required
def executive_pmo():
    projects=roadmap_accessible_projects()
    if request.method=="POST" and current_user()["role"] in ("admin","pm"):
        with db() as conn:
            conn.execute("INSERT OR IGNORE INTO programs(name,owner,description) VALUES(?,?,?)",
                         (request.form["name"].strip(),request.form.get("owner","").strip(),request.form.get("description","").strip()))
        return redirect(url_for("executive_pmo"))
    data=[]
    today=date.today()
    with db() as conn:
        programs=conn.execute("SELECT * FROM programs ORDER BY name").fetchall()
        for p in projects:
            tasks=[dict(x) for x in conn.execute("SELECT * FROM tasks WHERE project_id=?",(p["id"],))]
            risks=[dict(x) for x in conn.execute("SELECT * FROM risks WHERE project_id=? AND status<>'Stängd'",(p["id"],))]
            overdue=sum(1 for t in tasks if int(t.get("progress") or 0)<100 and parse_date(t.get("end_date")) and parse_date(t.get("end_date"))<today)
            high=sum(1 for r in risks if int(r.get("probability") or 0)*int(r.get("impact") or 0)>=12)
            progress=round(sum(int(t.get("progress") or 0) for t in tasks)/len(tasks)) if tasks else 0
            score=max(0,100-overdue*12-high*10-(100-progress)//5)
            rag="red" if score<60 else ("amber" if score<80 else "green")
            data.append({"p":p,"progress":progress,"overdue":overdue,"high":high,"score":score,"rag":rag})
    return render_template("executive_pmo.html",data=data,programs=programs)

@app.get("/portfolio/<int:portfolio_id>/dashboard")
@login_required
def portfolio_dashboard_v2(portfolio_id):
    with db() as conn:
        portfolio=conn.execute("SELECT * FROM portfolios WHERE id=?",(portfolio_id,)).fetchone()
        if not portfolio: abort(404)
        projects=conn.execute("""SELECT p.* FROM projects p JOIN portfolio_projects pp ON pp.project_id=p.id
                                 WHERE pp.portfolio_id=? ORDER BY p.name""",(portfolio_id,)).fetchall()
    return render_template("portfolio_dashboard_v2.html",portfolio=portfolio,projects=projects)

@app.route("/collaboration-hub",methods=["GET"])
@login_required
def collaboration_hub():
    projects=roadmap_accessible_projects(); u=current_user()
    ids=[p["id"] for p in projects]
    with db() as conn:
        mentions=conn.execute("""SELECT m.*,p.name project_name FROM mentions m JOIN projects p ON p.id=m.project_id
                                 WHERE m.mentioned_user_id=? ORDER BY m.id DESC LIMIT 30""",(u["id"],)).fetchall()
        if ids:
            ph=",".join("?" for _ in ids)
            meetings=conn.execute(f"SELECT m.*,p.name project_name FROM meetings m JOIN projects p ON p.id=m.project_id WHERE m.project_id IN ({ph}) ORDER BY meeting_date DESC LIMIT 30",ids).fetchall()
            actions=conn.execute(f"SELECT a.*,p.name project_name FROM action_items a JOIN projects p ON p.id=a.project_id WHERE a.project_id IN ({ph}) AND a.status<>'Done' ORDER BY a.due_date LIMIT 50",ids).fetchall()
        else: meetings=[]; actions=[]
    return render_template("collaboration_hub.html",projects=projects,mentions=mentions,meetings=meetings,actions=actions)

@app.post("/projects/<int:project_id>/mention")
@login_required
def create_mention(project_id):
    project_or_404(project_id,write=True)
    username=request.form.get("username","").strip().lstrip("@")
    with db() as conn:
        user=conn.execute("SELECT * FROM users WHERE username=? AND active=1",(username,)).fetchone()
        if not user:
            flash("Användaren hittades inte.","error")
        else:
            conn.execute("""INSERT INTO mentions(project_id,mentioned_user_id,source_type,message,created_at)
                            VALUES(?,?,?,?,?)""",(project_id,user["id"],request.form.get("source_type","message"),
                            request.form.get("message","").strip(),datetime.now().isoformat(timespec="seconds")))
            conn.execute("""INSERT INTO notifications(user_id,project_id,message,link,created_at)
                            VALUES(?,?,?,?,?)""",(user["id"],project_id,request.form.get("message","").strip(),
                            url_for("collaboration",project_id=project_id),datetime.now().isoformat(timespec="seconds")))
    return redirect(url_for("collaboration",project_id=project_id))

@app.post("/documents/<int:document_id>/version")
@login_required
def document_new_version(document_id):
    with db() as conn:
        doc=conn.execute("SELECT * FROM documents WHERE id=?",(document_id,)).fetchone()
        if not doc: abort(404)
    project_or_404(doc["project_id"],write=True)
    with db() as conn:
        n=conn.execute("SELECT COALESCE(MAX(version_no),0)+1 n FROM document_versions WHERE document_id=?",(document_id,)).fetchone()["n"]
        conn.execute("""INSERT INTO document_versions(document_id,version_no,content,created_by,created_at)
                        VALUES(?,?,?,?,?)""",(document_id,n,request.form.get("content",""),current_user()["id"],datetime.now().isoformat(timespec="seconds")))
        conn.execute("UPDATE documents SET body=?,updated_at=? WHERE id=?",(request.form.get("content",""),datetime.now().isoformat(timespec="seconds"),document_id))
    return redirect(url_for("collaboration",project_id=doc["project_id"]))

def project_assistant_summary(project_id):
    p=project_or_404(project_id)
    today=date.today()
    with db() as conn:
        tasks=[dict(r) for r in conn.execute("SELECT * FROM tasks WHERE project_id=?",(project_id,))]
        risks=[dict(r) for r in conn.execute("SELECT * FROM risks WHERE project_id=? AND status<>'Stängd'",(project_id,))]
        changes=[dict(r) for r in conn.execute("SELECT * FROM change_requests WHERE project_id=? AND status='Submitted'",(project_id,))]
        actions=[dict(r) for r in conn.execute("SELECT * FROM action_items WHERE project_id=? AND status<>'Done'",(project_id,))]
    overdue=[t for t in tasks if int(t.get("progress") or 0)<100 and parse_date(t.get("end_date")) and parse_date(t.get("end_date"))<today]
    blocked=[t for t in tasks if t.get("status")=="Blockerad"]
    high=[r for r in risks if int(r.get("probability") or 0)*int(r.get("impact") or 0)>=12]
    progress=round(sum(int(t.get("progress") or 0) for t in tasks)/len(tasks)) if tasks else 0
    lines=[f"{p['name']} är {progress}% färdigt."]
    if overdue: lines.append(f"{len(overdue)} aktivitet(er) är försenade.")
    if blocked: lines.append(f"{len(blocked)} aktivitet(er) är blockerade.")
    if high: lines.append(f"{len(high)} höga risker behöver uppmärksamhet.")
    if changes: lines.append(f"{len(changes)} change request(s) väntar på beslut.")
    if actions: lines.append(f"{len(actions)} öppna actions finns.")
    if not any((overdue,blocked,high,changes)): lines.append("Inga kritiska avvikelser identifierades av regelmotorn.")
    return {"project":p,"progress":progress,"overdue":overdue,"blocked":blocked,"high":high,"changes":changes,"actions":actions,"text":" ".join(lines)}

@app.route("/workspace")
@login_required
def workspace():
    projects=roadmap_accessible_projects()
    prefs={}
    with db() as conn:
        row=conn.execute("SELECT layout_json FROM dashboard_preferences WHERE user_id=?",(current_user()["id"],)).fetchone()
        if row:
            try: prefs=json.loads(row["layout_json"] or "{}")
            except: prefs={}
    return render_template("workspace.html",projects=projects,prefs=prefs)

@app.post("/workspace/preferences")
@login_required
def workspace_preferences():
    layout=request.form.get("layout_json","{}")
    try: json.loads(layout)
    except: layout="{}"
    with db() as conn:
        conn.execute("DELETE FROM dashboard_preferences WHERE user_id=?",(current_user()["id"],))
        conn.execute("INSERT INTO dashboard_preferences(user_id,layout_json) VALUES(?,?)",(current_user()["id"],layout))
    flash("Dashboard sparad.","success")
    return redirect(url_for("workspace"))

@app.get("/assistant/project/<int:project_id>")
@login_required
def assistant_project(project_id):
    summary=project_assistant_summary(project_id)
    return render_template("assistant_project.html",summary=summary)

@app.route("/report-center",methods=["GET","POST"])
@login_required
def report_center():
    if request.method=="POST":
        with db() as conn:
            conn.execute("INSERT INTO report_definitions(name,report_type,config_json,owner_user_id,created_at) VALUES(?,?,?,?,?)",
                         (request.form["name"].strip(),request.form.get("report_type","portfolio"),request.form.get("config_json","{}"),
                          current_user()["id"],datetime.now().isoformat(timespec="seconds")))
        return redirect(url_for("report_center"))
    with db() as conn:
        reports=conn.execute("SELECT * FROM report_definitions WHERE owner_user_id=? OR ?='admin' ORDER BY id DESC",(current_user()["id"],current_user()["role"])).fetchall()
    return render_template("report_center.html",reports=reports)

@app.route("/admin/integrations",methods=["GET","POST"])
@login_required
@role_required("admin")
def admin_integrations():
    with db() as conn:
        if request.method=="POST":
            conn.execute("UPDATE oidc_settings SET enabled=?,issuer=?,client_id=?,scopes=? WHERE id=1",
                         (1 if request.form.get("enabled")=="on" else 0,request.form.get("issuer","").strip(),
                          request.form.get("client_id","").strip(),request.form.get("scopes","openid profile email").strip()))
            flash("OIDC-konfiguration sparad. Aktivering kräver IdP/client secret i servermiljön.","success")
            return redirect(url_for("admin_integrations"))
        oidc=conn.execute("SELECT * FROM oidc_settings WHERE id=1").fetchone()
        webhooks=conn.execute("SELECT * FROM webhooks ORDER BY id DESC").fetchall()
        rules=conn.execute("SELECT * FROM automation_rules ORDER BY id DESC").fetchall()
    return render_template("admin_integrations.html",oidc=oidc,webhooks=webhooks,rules=rules)

@app.get("/api/v1/health/summary")
def api_health_summary():
    return jsonify(status="ok",version=APP_VERSION,features=["planning","control","resources","finance","pmo","collaboration","reports","assistant"])


def record_recent(entity_type,entity_id,label,url):
    u=current_user()
    if not u: return
    with db() as conn:
        conn.execute("DELETE FROM recent_items WHERE user_id=? AND entity_type=? AND entity_id=?",(u["id"],entity_type,entity_id))
        conn.execute("INSERT INTO recent_items(user_id,entity_type,entity_id,label,url,viewed_at) VALUES(?,?,?,?,?,?)",
                     (u["id"],entity_type,entity_id,label,url,datetime.now().isoformat(timespec="seconds")))
        conn.execute("DELETE FROM recent_items WHERE id IN (SELECT id FROM recent_items WHERE user_id=? ORDER BY viewed_at DESC LIMIT -1 OFFSET 20)",(u["id"],))

@app.get("/productivity")
@login_required
def productivity_hub():
    u=current_user()
    with db() as conn:
        fav=conn.execute("SELECT * FROM favorites WHERE user_id=? ORDER BY created_at DESC",(u["id"],)).fetchall()
        recent=conn.execute("SELECT * FROM recent_items WHERE user_id=? ORDER BY viewed_at DESC LIMIT 12",(u["id"],)).fetchall()
        views=conn.execute("SELECT * FROM saved_views WHERE user_id=? ORDER BY created_at DESC",(u["id"],)).fetchall()
    return render_template("productivity.html",favorites=fav,recent=recent,views=views)

@app.get("/global-search")
@login_required
def global_search():
    q=request.args.get("q","").strip(); results=[]
    if q:
        like=f"%{q}%"; projects=roadmap_accessible_projects(); ids=[p["id"] for p in projects]
        with db() as conn:
            for p in projects:
                if q.lower() in (p["name"]+" "+(p["customer"] or "")).lower():
                    results.append({"type":"Projekt","label":p["name"],"detail":p["customer"],"url":url_for("project_cockpit",project_id=p["id"])})
            if ids:
                ph=",".join("?" for _ in ids)
                rows=conn.execute(f"SELECT t.*,p.name project_name FROM tasks t JOIN projects p ON p.id=t.project_id WHERE t.project_id IN ({ph}) AND (t.title LIKE ? OR t.wbs LIKE ?) LIMIT 50",(*ids,like,like)).fetchall()
                for t in rows: results.append({"type":"Aktivitet","label":t["title"],"detail":f"{t['project_name']} · {t['wbs']}","url":url_for("edit_task",project_id=t["project_id"],task_id=t["id"])})
    return render_template("global_search.html",q=q,results=results)

@app.post("/favorites/toggle")
@login_required
def toggle_favorite():
    et=request.form["entity_type"]; eid=int(request.form["entity_id"]); u=current_user()
    with db() as conn:
        row=conn.execute("SELECT 1 FROM favorites WHERE user_id=? AND entity_type=? AND entity_id=?",(u["id"],et,eid)).fetchone()
        if row: conn.execute("DELETE FROM favorites WHERE user_id=? AND entity_type=? AND entity_id=?",(u["id"],et,eid))
        else: conn.execute("INSERT INTO favorites(user_id,entity_type,entity_id,created_at) VALUES(?,?,?,?)",(u["id"],et,eid,datetime.now().isoformat(timespec="seconds")))
    return redirect(request.referrer or url_for("productivity_hub"))

@app.post("/projects/<int:project_id>/tasks/bulk")
@login_required
def tasks_bulk_update(project_id):
    project_or_404(project_id,write=True)
    ids=[int(x) for x in request.form.getlist("task_id") if x.isdigit()]
    if ids:
        status=request.form.get("status",""); priority=request.form.get("priority","")
        with db() as conn:
            for tid in ids:
                if status: conn.execute("UPDATE tasks SET status=? WHERE id=? AND project_id=?",(status,tid,project_id))
                if priority: conn.execute("UPDATE tasks SET priority=? WHERE id=? AND project_id=?",(priority,tid,project_id))
        audit(project_id,"task",None,"bulk_updated",f"{len(ids)} activities")
    return redirect(request.referrer or url_for("project_wbs",project_id=project_id))

@app.get("/projects/<int:project_id>/timeline")
@login_required
def project_timeline(project_id):
    p=project_or_404(project_id); record_recent("project",project_id,p["name"],url_for("project_cockpit",project_id=project_id))
    with db() as conn:
        rows=conn.execute("SELECT a.*,u.display_name FROM audit_log a LEFT JOIN users u ON u.id=a.user_id WHERE a.project_id=? ORDER BY a.id DESC LIMIT 100",(project_id,)).fetchall()
    return render_template("project_timeline.html",project=p,rows=rows)


def planning_network(project_id):
    with db() as conn:
        tasks=[dict(r) for r in conn.execute("SELECT * FROM tasks WHERE project_id=?",(project_id,))]
        links=[dict(r) for r in conn.execute("SELECT * FROM task_links WHERE project_id=?",(project_id,))]
    by={int(t["id"]):t for t in tasks}; preds={k:[] for k in by}; succ={k:[] for k in by}
    for l in links:
        a=int(l["predecessor_id"]); b=int(l["successor_id"])
        if a in by and b in by: preds[b].append(l); succ[a].append(l)
    indeg={k:len(preds[k]) for k in by}; q=[k for k,v in indeg.items() if v==0]; order=[]
    while q:
        n=q.pop(0); order.append(n)
        for l in succ[n]:
            b=int(l["successor_id"]); indeg[b]-=1
            if indeg[b]==0:q.append(b)
    cycle=len(order)!=len(by)
    if cycle: order=list(by)
    es={k:0 for k in by}; ef={}
    for k in order:
        dur=max(1,int(by[k].get("duration_days") or 1)); start=0
        for l in preds[k]:
            a=int(l["predecessor_id"]); lag=int(l.get("lag_days") or 0); typ=l.get("link_type") or "FS"
            pd=max(1,int(by[a].get("duration_days") or 1))
            if typ=="FS": cand=ef.get(a,pd)+lag
            elif typ=="SS": cand=es.get(a,0)+lag
            elif typ=="FF": cand=ef.get(a,pd)+lag-dur
            else: cand=es.get(a,0)+lag-dur
            start=max(start,cand)
        es[k]=max(0,start); ef[k]=es[k]+dur
    finish=max(ef.values(),default=0); lf={k:finish for k in by}; ls={}
    for k in reversed(order):
        dur=max(1,int(by[k].get("duration_days") or 1)); latest=finish-dur
        if succ[k]:
            vals=[]
            for l in succ[k]:
                b=int(l["successor_id"]); lag=int(l.get("lag_days") or 0); typ=l.get("link_type") or "FS"; bd=max(1,int(by[b].get("duration_days") or 1))
                if typ=="FS": vals.append(ls.get(b,finish-bd)-lag-dur)
                elif typ=="SS": vals.append(ls.get(b,finish-bd)-lag)
                elif typ=="FF": vals.append(lf.get(b,finish)-lag-dur)
                else: vals.append(lf.get(b,finish)-lag)
            latest=min(vals)
        ls[k]=max(0,latest); lf[k]=ls[k]+dur
    rows=[]
    for k in order:
        t=dict(by[k]); t.update(es=es[k],ef=ef[k],ls=ls[k],lf=lf[k],slack=max(0,ls[k]-es[k]),critical=(ls[k]-es[k])==0); rows.append(t)
    return rows,links,cycle,finish

@app.get("/projects/<int:project_id>/advanced-planning")
@login_required
def advanced_planning(project_id):
    p=project_or_404(project_id); rows,links,cycle,finish=planning_network(project_id)
    with db() as conn: scenarios=conn.execute("SELECT * FROM planning_scenarios WHERE project_id=? ORDER BY id DESC",(project_id,)).fetchall()
    return render_template("advanced_planning.html",project=p,rows=rows,links=links,cycle=cycle,finish=finish,scenarios=scenarios)

@app.post("/projects/<int:project_id>/dependencies/new")
@login_required
def advanced_dependency_new(project_id):
    project_or_404(project_id,write=True); a=int(request.form["predecessor_id"]); b=int(request.form["successor_id"])
    if a==b: abort(400,"En aktivitet kan inte bero på sig själv")
    with db() as conn:
        conn.execute("INSERT INTO task_links(project_id,predecessor_id,successor_id,link_type,lag_days) VALUES(?,?,?,?,?)",(project_id,a,b,request.form.get("link_type","FS"),int(request.form.get("lag_days","0") or 0)))
    audit(project_id,"dependency",None,"created",f"{a}->{b}")
    return redirect(url_for("advanced_planning",project_id=project_id))

@app.post("/projects/<int:project_id>/scenario")
@login_required
def planning_scenario_save(project_id):
    project_or_404(project_id,write=True)
    with db() as conn:
        tasks=[dict(r) for r in conn.execute("SELECT * FROM tasks WHERE project_id=?",(project_id,))]; links=[dict(r) for r in conn.execute("SELECT * FROM task_links WHERE project_id=?",(project_id,))]
        conn.execute("INSERT INTO planning_scenarios(project_id,name,snapshot_json,created_by,created_at) VALUES(?,?,?,?,?)",(project_id,request.form["name"].strip(),json.dumps({"tasks":tasks,"links":links},ensure_ascii=False),current_user()["id"],datetime.now().isoformat(timespec="seconds")))
    return redirect(url_for("advanced_planning",project_id=project_id))

@app.get("/projects/<int:project_id>/baseline-variance")
@login_required
def baseline_variance(project_id):
    p=project_or_404(project_id)
    with db() as conn:
        current=[dict(r) for r in conn.execute("SELECT * FROM tasks WHERE project_id=? ORDER BY wbs,id",(project_id,))]
        base=conn.execute("SELECT * FROM baselines WHERE project_id=? ORDER BY id DESC LIMIT 1",(project_id,)).fetchone()
    snap=json.loads(base["snapshot_json"]) if base else []
    if isinstance(snap,dict): snap=snap.get("tasks",[])
    bm={str(x.get("id")):x for x in snap if isinstance(x,dict)}
    rows=[]
    for t in current:
        b=bm.get(str(t["id"])) or next((x for x in snap if isinstance(x,dict) and x.get("wbs")==t.get("wbs")),{})
        delay=0
        if b and parse_date(b.get("end_date")) and parse_date(t.get("end_date")): delay=(parse_date(t["end_date"])-parse_date(b["end_date"])).days
        rows.append({"task":t,"baseline":b,"delay":delay})
    return render_template("baseline_variance.html",project=p,baseline=base,rows=rows)


def evaluate_automation_rule(rule):
    cfg=json.loads(rule["config_json"] or "{}")
    pid=rule["project_id"]; trigger=rule["trigger_type"]; action=rule["action_type"]; matched=[]
    with db() as conn:
        if trigger=="task.overdue" and pid:
            for t in conn.execute("SELECT * FROM tasks WHERE project_id=? AND progress<100 AND end_date<>''",(pid,)).fetchall():
                if parse_date(t["end_date"]) and parse_date(t["end_date"])<date.today(): matched.append(dict(t))
        elif trigger=="risk.high" and pid:
            matched=[dict(r) for r in conn.execute("SELECT * FROM risks WHERE project_id=? AND probability*impact>=? AND status<>'Stängd'",(pid,int(cfg.get("score",12)))).fetchall()]
        elif trigger=="milestone.due" and pid:
            days=int(cfg.get("days",7)); until=date.today()+timedelta(days=days)
            for t in conn.execute("SELECT * FROM tasks WHERE project_id=? AND milestone=1 AND progress<100",(pid,)).fetchall():
                d=parse_date(t["end_date"]); matched += [dict(t)] if d and date.today()<=d<=until else []
        if matched and action=="notify.project.members" and pid:
            members=conn.execute("SELECT user_id FROM project_members WHERE project_id=?",(pid,)).fetchall()
            msg=cfg.get("message") or f"Automation: {rule['name']} matched {len(matched)} item(s)."
            for m in members: conn.execute("INSERT INTO notifications(user_id,project_id,message,link,created_at) VALUES(?,?,?,?,?)",(m["user_id"],pid,msg,url_for("project_cockpit",project_id=pid),datetime.now().isoformat(timespec="seconds")))
    return matched

def run_automation_rules(project_id=None):
    with db() as conn:
        if project_id: rules=conn.execute("SELECT * FROM automation_rules WHERE active=1 AND (project_id=? OR project_id IS NULL)",(project_id,)).fetchall()
        else: rules=conn.execute("SELECT * FROM automation_rules WHERE active=1").fetchall()
    results=[]
    for rule in rules:
        started=datetime.now().isoformat(timespec="seconds")
        try:
            matched=evaluate_automation_rule(rule); status="Matched" if matched else "No match"; detail=f"{len(matched)} item(s)"
        except Exception as ex: status="Failed"; detail=str(ex)
        with db() as conn: conn.execute("INSERT INTO automation_executions(rule_id,project_id,status,details,started_at,finished_at) VALUES(?,?,?,?,?,?)",(rule["id"],rule["project_id"],status,detail,started,datetime.now().isoformat(timespec="seconds")))
        results.append((rule,status,detail))
    return results

@app.route("/automation-center",methods=["GET","POST"])
@login_required
def automation_center():
    projects=roadmap_accessible_projects()
    if request.method=="POST":
        pid=int(request.form["project_id"]) if request.form.get("project_id") else None
        if pid: project_or_404(pid,manager=True)
        with db() as conn: conn.execute("INSERT INTO automation_rules(project_id,name,trigger_type,action_type,config_json,active) VALUES(?,?,?,?,?,1)",(pid,request.form["name"].strip(),request.form["trigger_type"],request.form["action_type"],request.form.get("config_json","{}")))
        return redirect(url_for("automation_center"))
    with db() as conn:
        rules=conn.execute("SELECT * FROM automation_rules ORDER BY id DESC").fetchall(); history=conn.execute("SELECT e.*,r.name rule_name FROM automation_executions e LEFT JOIN automation_rules r ON r.id=e.rule_id ORDER BY e.id DESC LIMIT 100").fetchall()
    return render_template("automation_center.html",projects=projects,rules=rules,history=history)

@app.post("/automation/run")
@login_required
@role_required("admin","pm")
def automation_run_now():
    results=run_automation_rules(); flash(f"Automation körd: {len(results)} regel/regler.","success"); return redirect(url_for("automation_center"))

@app.post("/internal/automation/run")
def automation_internal_run():
    token=request.headers.get("X-Automation-Token","")
    expected=os.getenv("PROJECT_PLAN_AUTOMATION_TOKEN","")
    if not expected or not hmac.compare_digest(token,expected): abort(401)
    return jsonify(results=[{"rule":r[0]["name"],"status":r[1],"details":r[2]} for r in run_automation_rules()])


def project_report_payload(project_id):
    s=project_assistant_summary(project_id); p=s["project"]
    with db() as conn:
        milestones=[dict(r) for r in conn.execute("SELECT * FROM tasks WHERE project_id=? AND milestone=1 ORDER BY end_date",(project_id,))]
        finance=conn.execute("SELECT * FROM project_finance WHERE project_id=?",(project_id,)).fetchone()
        costs=conn.execute("SELECT COALESCE(SUM(actual),0) actual,COALESCE(SUM(planned),0) planned FROM project_costs WHERE project_id=?",(project_id,)).fetchone()
    return {"summary":s,"project":p,"milestones":milestones,"finance":dict(finance) if finance else {},"costs":dict(costs)}

@app.get("/projects/<int:project_id>/report.pdf")
@login_required
def project_report_pdf(project_id):
    payload=project_report_payload(project_id)
    from reportlab.lib.pagesizes import A4
    from reportlab.pdfgen import canvas
    b=BytesIO(); c=canvas.Canvas(b,pagesize=A4); w,h=A4; y=h-55
    c.setFont("Helvetica-Bold",18); c.drawString(45,y,f"Project Status – {payload['project']['name']}"); y-=30
    c.setFont("Helvetica",10); c.drawString(45,y,f"Generated {datetime.now().strftime('%Y-%m-%d %H:%M')} · Project Planer {APP_VERSION}"); y-=28
    text=c.beginText(45,y); text.setFont("Helvetica",11)
    for line in [payload["summary"]["text"],f"Progress: {payload['summary']['progress']}%",f"Overdue: {len(payload['summary']['overdue'])}",f"High risks: {len(payload['summary']['high'])}"]:
        for part in [line[i:i+95] for i in range(0,len(line),95)]: text.textLine(part)
    c.drawText(text); y=text.getY()-20; c.setFont("Helvetica-Bold",12); c.drawString(45,y,"Milestones"); y-=18; c.setFont("Helvetica",10)
    for m in payload["milestones"][:18]: c.drawString(55,y,f"{m.get('end_date','')}  {m.get('title','')}  {m.get('progress',0)}%"); y-=15
    c.showPage(); c.save(); b.seek(0)
    return send_file(b,mimetype="application/pdf",as_attachment=True,download_name=f"{secure_filename(payload['project']['name'])}-status.pdf")

@app.get("/projects/<int:project_id>/report.pptx")
@login_required
def project_report_pptx(project_id):
    payload=project_report_payload(project_id); from pptx import Presentation
    prs=Presentation(); slide=prs.slides.add_slide(prs.slide_layouts[1]); slide.shapes.title.text=payload["project"]["name"]+" – Project Status"; slide.placeholders[1].text=payload["summary"]["text"]
    slide=prs.slides.add_slide(prs.slide_layouts[1]); slide.shapes.title.text="Key metrics"; slide.placeholders[1].text=f"Progress: {payload['summary']['progress']}%\nOverdue: {len(payload['summary']['overdue'])}\nBlocked: {len(payload['summary']['blocked'])}\nHigh risks: {len(payload['summary']['high'])}"
    slide=prs.slides.add_slide(prs.slide_layouts[1]); slide.shapes.title.text="Milestones"; slide.placeholders[1].text="\n".join(f"{m.get('end_date','')} – {m.get('title','')} ({m.get('progress',0)}%)" for m in payload["milestones"][:12]) or "No milestones"
    b=BytesIO(); prs.save(b); b.seek(0); return send_file(b,mimetype="application/vnd.openxmlformats-officedocument.presentationml.presentation",as_attachment=True,download_name=f"{secure_filename(payload['project']['name'])}-status.pptx")

@app.get("/executive-reporting")
@login_required
def executive_reporting():
    rows=[]
    for p in roadmap_accessible_projects():
        s=project_assistant_summary(p["id"]); rag="red" if s["overdue"] or len(s["high"])>=2 else ("amber" if s["blocked"] or s["high"] else "green")
        rows.append({"p":p,"s":s,"rag":rag})
    return render_template("executive_reporting.html",rows=rows)

@app.route("/scheduled-reports",methods=["GET","POST"])
@login_required
def scheduled_reports():
    if request.method=="POST":
        pid=int(request.form["project_id"]) if request.form.get("project_id") else None
        if pid: project_or_404(pid,manager=True)
        with db() as conn: conn.execute("INSERT INTO scheduled_reports(project_id,name,format,cadence,recipient,created_by,created_at) VALUES(?,?,?,?,?,?,?)",(pid,request.form["name"],request.form.get("format","pdf"),request.form.get("cadence","weekly"),request.form.get("recipient",""),current_user()["id"],datetime.now().isoformat(timespec="seconds")))
        return redirect(url_for("scheduled_reports"))
    with db() as conn: rows=conn.execute("SELECT * FROM scheduled_reports ORDER BY id DESC").fetchall()
    return render_template("scheduled_reports.html",rows=rows,projects=roadmap_accessible_projects())


def api_auth_scope(scope="read"):
    def deco(fn):
        @wraps(fn)
        def wrapper(*args,**kwargs):
            auth=request.headers.get("Authorization","")
            if not auth.startswith("Bearer "): return jsonify(error="unauthorized"),401
            raw=auth[7:]; h=hashlib.sha256(raw.encode()).hexdigest()
            with db() as conn: key=conn.execute("SELECT * FROM api_keys WHERE key_hash=? AND active=1",(h,)).fetchone()
            if not key: return jsonify(error="unauthorized"),401
            scopes=set((key["scopes"] or "read").split());
            if scope not in scopes and "admin" not in scopes: return jsonify(error="insufficient_scope"),403
            if key["expires_at"] and parse_date(key["expires_at"]) and parse_date(key["expires_at"])<date.today(): return jsonify(error="key_expired"),401
            request.api_user=key["user_id"]; conn.execute("UPDATE api_keys SET last_used=? WHERE id=?",(datetime.now().isoformat(timespec="seconds"),key["id"]))
            return fn(*args,**kwargs)
        return wrapper
    return deco

@app.route("/api/v2/projects",methods=["GET","POST"])
@api_auth_scope("read")
def api_v2_projects():
    uid=request.api_user
    with db() as conn:
        u=conn.execute("SELECT * FROM users WHERE id=?",(uid,)).fetchone()
        if request.method=="POST":
            scopes=request.headers.get("X-API-Scope-Check","")
            # write authorization is revalidated against token below
            auth=request.headers.get("Authorization","")[7:]; kh=hashlib.sha256(auth.encode()).hexdigest(); key=conn.execute("SELECT * FROM api_keys WHERE key_hash=?",(kh,)).fetchone()
            if "write" not in set((key["scopes"] or "").split()) and "admin" not in set((key["scopes"] or "").split()): return jsonify(error="insufficient_scope"),403
            data=request.get_json(silent=True) or {}; name=(data.get("name") or "").strip()
            if not name:return jsonify(error="name_required"),400
            cur=conn.execute("INSERT INTO projects(name,customer,project_manager,description,created_by,created_at) VALUES(?,?,?,?,?,?)",(name,data.get("customer",""),data.get("project_manager",""),data.get("description",""),uid,datetime.now().isoformat(timespec="seconds")))
            pid=cur.lastrowid; conn.execute("INSERT OR IGNORE INTO project_members(project_id,user_id,project_role,added_at,added_by) VALUES(?,?,?,?,?)",(pid,uid,"pm",datetime.now().isoformat(timespec="seconds"),uid)); return jsonify(id=pid,name=name),201
        if u["role"]=="admin": rows=conn.execute("SELECT * FROM projects ORDER BY id DESC").fetchall()
        else: rows=conn.execute("SELECT p.* FROM projects p JOIN project_members pm ON pm.project_id=p.id WHERE pm.user_id=? ORDER BY p.id DESC",(uid,)).fetchall()
    return jsonify([dict(r) for r in rows])

@app.route("/api/v2/projects/<int:project_id>/tasks",methods=["GET","POST"])
@api_auth_scope("read")
def api_v2_tasks(project_id):
    uid=request.api_user
    with db() as conn:
        u=conn.execute("SELECT * FROM users WHERE id=?",(uid,)).fetchone(); allowed=u["role"]=="admin" or conn.execute("SELECT 1 FROM project_members WHERE project_id=? AND user_id=?",(project_id,uid)).fetchone()
        if not allowed:return jsonify(error="forbidden"),403
        if request.method=="POST":
            auth=request.headers.get("Authorization","")[7:]; key=conn.execute("SELECT * FROM api_keys WHERE key_hash=?",(hashlib.sha256(auth.encode()).hexdigest(),)).fetchone(); scopes=set((key["scopes"] or "").split())
            if "write" not in scopes and "admin" not in scopes:return jsonify(error="insufficient_scope"),403
            d=request.get_json(silent=True) or {}; title=(d.get("title") or "").strip()
            if not title:return jsonify(error="title_required"),400
            cur=conn.execute("INSERT INTO tasks(project_id,wbs,title,owner,start_date,end_date,status,priority,progress) VALUES(?,?,?,?,?,?,?,?,?)",(project_id,d.get("wbs",""),title,d.get("owner",""),d.get("start_date",""),d.get("end_date",""),d.get("status","Ej startad"),d.get("priority","Normal"),int(d.get("progress",0))))
            return jsonify(id=cur.lastrowid,title=title),201
        rows=conn.execute("SELECT * FROM tasks WHERE project_id=? ORDER BY wbs,id",(project_id,)).fetchall()
    return jsonify([dict(r) for r in rows])

@app.route("/integration-center",methods=["GET","POST"])
@login_required
@role_required("admin")
def integration_center_v35():
    providers=["Azure DevOps","GitHub","Jira","Microsoft Teams","SMTP"]
    if request.method=="POST":
        provider=request.form["provider"]
        with db() as conn: conn.execute("INSERT INTO integration_settings(provider,enabled,base_url,config_json,updated_at) VALUES(?,?,?,?,?) ON CONFLICT(provider) DO UPDATE SET enabled=excluded.enabled,base_url=excluded.base_url,config_json=excluded.config_json,updated_at=excluded.updated_at",(provider,1 if request.form.get("enabled") else 0,request.form.get("base_url",""),request.form.get("config_json","{}"),datetime.now().isoformat(timespec="seconds")))
        return redirect(url_for("integration_center_v35"))
    with db() as conn: settings={r["provider"]:r for r in conn.execute("SELECT * FROM integration_settings").fetchall()}
    return render_template("integration_center_v35.html",providers=providers,settings=settings)


@app.get("/security-center")
@login_required
@role_required("admin")
def security_center():
    with db() as conn:
        logins=conn.execute("SELECT * FROM login_history ORDER BY id DESC LIMIT 100").fetchall(); events=conn.execute("SELECT * FROM security_events ORDER BY id DESC LIMIT 100").fetchall(); sessions=conn.execute("SELECT s.*,u.username,u.display_name FROM user_sessions s JOIN users u ON u.id=s.user_id WHERE s.revoked_at='' ORDER BY s.last_seen DESC").fetchall()
    return render_template("security_center.html",logins=logins,events=events,sessions=sessions)

@app.post("/projects/<int:project_id>/archive")
@login_required
def project_archive(project_id):
    project_or_404(project_id,manager=True)
    with db() as conn: conn.execute("UPDATE projects SET archived_at=? WHERE id=?",(datetime.now().isoformat(timespec="seconds"),project_id))
    audit(project_id,"project",project_id,"archived",""); return redirect(url_for("index"))

@app.post("/projects/<int:project_id>/trash")
@login_required
def project_trash(project_id):
    project_or_404(project_id,manager=True)
    with db() as conn: conn.execute("UPDATE projects SET deleted_at=? WHERE id=?",(datetime.now().isoformat(timespec="seconds"),project_id))
    audit(project_id,"project",project_id,"trashed",""); return redirect(url_for("index"))

@app.get("/admin/trash")
@login_required
@role_required("admin")
def admin_trash():
    with db() as conn: rows=conn.execute("SELECT * FROM projects WHERE deleted_at<>'' ORDER BY deleted_at DESC").fetchall()
    return render_template("trash.html",rows=rows)

@app.post("/admin/trash/<int:project_id>/restore")
@login_required
@role_required("admin")
def admin_trash_restore(project_id):
    with db() as conn: conn.execute("UPDATE projects SET deleted_at='' WHERE id=?",(project_id,))
    return redirect(url_for("admin_trash"))

@app.get("/admin/backup/download-latest")
@login_required
@role_required("admin")
def download_live_database():
    if not DB_PATH.exists(): abort(404)
    stamp=datetime.now().strftime("%Y%m%d-%H%M%S")
    return send_file(DB_PATH,as_attachment=True,download_name=f"projectplan-{stamp}.db")

@app.after_request
def enterprise_security_headers(response):
    response.headers.setdefault("X-Content-Type-Options","nosniff"); response.headers.setdefault("X-Frame-Options","SAMEORIGIN"); response.headers.setdefault("Referrer-Policy","strict-origin-when-cross-origin"); response.headers.setdefault("Permissions-Policy","camera=(), microphone=(), geolocation=()")
    if request.is_secure: response.headers.setdefault("Strict-Transport-Security","max-age=31536000; includeSubDomains")
    return response


@app.route("/projects/<int:project_id>/delivery",methods=["GET","POST"])
@login_required
def delivery_center(project_id):
    p=project_or_404(project_id)
    if request.method=="POST":
        project_or_404(project_id,write=True); kind=request.form["kind"]
        with db() as conn:
            if kind=="environment": conn.execute("INSERT OR IGNORE INTO project_environments(project_id,name,base_url,owner,status,notes) VALUES(?,?,?,?,?,?)",(project_id,request.form["name"],request.form.get("base_url",""),request.form.get("owner",""),request.form.get("status","Ready"),request.form.get("notes","")))
            elif kind=="deliverable": conn.execute("INSERT INTO deliverables(project_id,name,owner,due_date,status,acceptance_criteria) VALUES(?,?,?,?,?,?)",(project_id,request.form["name"],request.form.get("owner",""),request.form.get("due_date",""),request.form.get("status","Planned"),request.form.get("acceptance_criteria","")))
            elif kind=="interface": conn.execute("INSERT INTO interface_register(project_id,name,source_system,target_system,protocol,owner,status,notes) VALUES(?,?,?,?,?,?,?,?)",(project_id,request.form["name"],request.form.get("source_system",""),request.form.get("target_system",""),request.form.get("protocol",""),request.form.get("owner",""),request.form.get("status","Design"),request.form.get("notes","")))
            elif kind=="test": conn.execute("INSERT INTO test_cycles(project_id,name,test_type,start_date,end_date,status) VALUES(?,?,?,?,?,?)",(project_id,request.form["name"],request.form.get("test_type","SIT"),request.form.get("start_date",""),request.form.get("end_date",""),request.form.get("status","Planned")))
            elif kind=="requirement": conn.execute("INSERT INTO requirements_traceability(project_id,requirement_id,requirement,status) VALUES(?,?,?,?)",(project_id,request.form["requirement_id"],request.form["requirement"],request.form.get("status","Open")))
            elif kind=="cutover": conn.execute("INSERT INTO cutover_items(project_id,sequence_no,title,owner,planned_at,status,rollback_step) VALUES(?,?,?,?,?,?,?)",(project_id,int(request.form.get("sequence_no","0") or 0),request.form["title"],request.form.get("owner",""),request.form.get("planned_at",""),request.form.get("status","Planned"),request.form.get("rollback_step","")))
        audit(project_id,"delivery",None,"created",kind); return redirect(url_for("delivery_center",project_id=project_id))
    with db() as conn:
        envs=conn.execute("SELECT * FROM project_environments WHERE project_id=? ORDER BY name",(project_id,)).fetchall(); deliverables=conn.execute("SELECT * FROM deliverables WHERE project_id=? ORDER BY due_date",(project_id,)).fetchall(); interfaces=conn.execute("SELECT * FROM interface_register WHERE project_id=? ORDER BY name",(project_id,)).fetchall(); tests=conn.execute("SELECT * FROM test_cycles WHERE project_id=? ORDER BY start_date",(project_id,)).fetchall(); reqs=conn.execute("SELECT * FROM requirements_traceability WHERE project_id=? ORDER BY requirement_id",(project_id,)).fetchall(); cutover=conn.execute("SELECT * FROM cutover_items WHERE project_id=? ORDER BY sequence_no,id",(project_id,)).fetchall()
    return render_template("delivery_center.html",project=p,envs=envs,deliverables=deliverables,interfaces=interfaces,tests=tests,reqs=reqs,cutover=cutover)


def health_snapshot(project_id):
    s=project_assistant_summary(project_id); score=max(0,min(100,100-len(s["overdue"])*8-len(s["blocked"])*7-len(s["high"])*10-len(s["changes"])*3))
    return {"progress":s["progress"],"overdue":len(s["overdue"]),"blocked":len(s["blocked"]),"high":len(s["high"]),"changes":len(s["changes"]),"score":score,"summary":s}

def assistant_answer(project_id,question):
    q=question.lower(); h=health_snapshot(project_id); s=h["summary"]; p=s["project"]
    if any(x in q for x in ["sen","försen","late","delay"]):
        items=s["overdue"]; return (f"{len(items)} aktivitet(er) är försenade: "+", ".join(x["title"] for x in items[:10])) if items else "Inga försenade aktiviteter hittades."
    if any(x in q for x in ["block","go-live","golive"]):
        items=s["blocked"]+s["overdue"]; return ("De viktigaste blockerarna är: "+", ".join(dict.fromkeys(x["title"] for x in items[:10]))) if items else "Inga direkta blockerare identifierades."
    if "risk" in q:
        items=s["high"]; return (f"{len(items)} höga risker: "+", ".join(x["title"] for x in items[:10])) if items else "Inga höga risker identifierades."
    if any(x in q for x in ["ändrat","changed","förra veckan","last week"]):
        with db() as conn: changes=conn.execute("SELECT * FROM audit_log WHERE project_id=? AND created_at>=? ORDER BY id DESC LIMIT 30",(project_id,(datetime.now()-timedelta(days=7)).isoformat(timespec="seconds"))).fetchall()
        return f"{len(changes)} ändringar senaste 7 dagarna. "+"; ".join(f"{x['entity_type']} {x['action']}" for x in changes[:12])
    if any(x in q for x in ["hälsa","status","health"]): return f"{p['name']} har health score {h['score']}/100 och progress {h['progress']}%. {s['text']}"
    return s["text"]+" Fråga gärna om risker, förseningar, blockerare, status eller vad som ändrats senaste veckan."

@app.route("/intelligence",methods=["GET","POST"])
@login_required
def intelligence_center():
    projects=roadmap_accessible_projects(); pid=request.values.get("project_id",type=int) or (projects[0]["id"] if projects else None); answer=None; question=""
    if pid: project_or_404(pid)
    if request.method=="POST" and pid:
        question=request.form.get("question","").strip(); answer=assistant_answer(pid,question)
        with db() as conn: conn.execute("INSERT INTO assistant_queries(user_id,project_id,question,answer,created_at) VALUES(?,?,?,?,?)",(current_user()["id"],pid,question,answer,datetime.now().isoformat(timespec="seconds")))
    hist=[]
    if pid:
        with db() as conn: hist=conn.execute("SELECT * FROM assistant_queries WHERE user_id=? AND project_id=? ORDER BY id DESC LIMIT 10",(current_user()["id"],pid)).fetchall()
    return render_template("intelligence.html",projects=projects,project_id=pid,answer=answer,question=question,history=hist)

@app.post("/projects/<int:project_id>/health-snapshot")
@login_required
def create_health_snapshot(project_id):
    project_or_404(project_id,write=True); h=health_snapshot(project_id)
    with db() as conn: conn.execute("INSERT INTO project_health_snapshots(project_id,snapshot_date,progress,overdue_count,blocked_count,high_risk_count,open_change_count,health_score,details_json) VALUES(?,?,?,?,?,?,?,?,?) ON CONFLICT(project_id,snapshot_date) DO UPDATE SET progress=excluded.progress,overdue_count=excluded.overdue_count,blocked_count=excluded.blocked_count,high_risk_count=excluded.high_risk_count,open_change_count=excluded.open_change_count,health_score=excluded.health_score,details_json=excluded.details_json",(project_id,date.today().isoformat(),h["progress"],h["overdue"],h["blocked"],h["high"],h["changes"],h["score"],json.dumps({"version":APP_VERSION})))
    return redirect(url_for("intelligence_center",project_id=project_id))

@app.get("/projects/<int:project_id>/health-trend")
@login_required
def health_trend(project_id):
    p=project_or_404(project_id)
    with db() as conn: rows=conn.execute("SELECT * FROM project_health_snapshots WHERE project_id=? ORDER BY snapshot_date",(project_id,)).fetchall()
    return render_template("health_trend.html",project=p,rows=rows)

@app.post("/projects/<int:project_id>/meeting-to-actions")
@login_required
def meeting_to_actions(project_id):
    project_or_404(project_id,write=True); notes=request.form.get("notes",""); created=0
    with db() as conn:
        for line in notes.splitlines():
            raw=line.strip()
            if raw.upper().startswith("ACTION:"):
                body=raw.split(":",1)[1].strip(); owner=""
                if "@" in body:
                    body,owner=body.rsplit("@",1); body=body.strip(); owner=owner.strip()
                if body: conn.execute("INSERT INTO action_items(project_id,title,owner,status) VALUES(?,?,?,'Open')",(project_id,body,owner)); created+=1
            elif raw.upper().startswith("DECISION:"):
                body=raw.split(":",1)[1].strip()
                if body: conn.execute("INSERT INTO decisions(project_id,title,decision,decided_by,decision_date,owner,decided_at) VALUES(?,?,?,?,?,?,?)",(project_id,body,body,current_user()["display_name"],date.today().isoformat(),current_user()["display_name"],datetime.now().isoformat(timespec="seconds"))); created+=1
    flash(f"Skapade {created} actions/beslut från mötesanteckningar.","success"); return redirect(url_for("collaboration",project_id=project_id))

if __name__=="__main__":
    init_db()
    app.run(host="0.0.0.0",port=int(os.getenv("PORT","8080")),debug=False)
