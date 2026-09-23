import tempfile
import unittest
from unittest.mock import patch
from pathlib import Path

from openpyxl import load_workbook
from werkzeug.security import generate_password_hash

from app import create_app, registry_rows, score


class AppTestCase(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        root = Path(self.temp.name)
        self.app = create_app({"TESTING": True, "DATABASE": str(root / "test.db"), "UPLOAD_FOLDER": str(root / "uploads"), "SECRET_KEY": "test", "DEEPSEEK_API_KEY": "", "PUBLIC_DEMO_MODE": False})
        self.client = self.app.test_client()

    def tearDown(self):
        self.temp.cleanup()

    def csrf(self):
        with self.client.session_transaction() as session:
            return session["csrf_token"]

    def login(self, username="admin", password="admin123"):
        with self.client.session_transaction() as session:
            session.clear()
        self.client.get("/login")
        return self.client.post("/login", data={"profile": "organizer", "username": username,
                                                "password": password, "csrf_token": self.csrf()}, follow_redirects=True)

    def add_user(self, username, role):
        with self.app.app_context():
            db = self.app.get_db()
            db.execute("INSERT INTO users(username,password_hash,role) VALUES(?,?,?)", (username, generate_password_hash("password1", method="pbkdf2:sha256"), role))
            db.commit()

    def achievement_data(self):
        return {"csrf_token": self.csrf(), "child_name": "Тестов Тимур Павлович", "school": "Гимназия Тест", "class_name": "5В", "contest_name": "Конференция юных исследователей", "event_date": "2026-05-10", "level": "all_russian", "result": "first", "direction": "", "mentor_name": "Наставников Николай", "mentor_position": "учитель", "social_notes": "Тестовая запись"}

    def opportunity_data(self, direction="creative", title="Фестиваль талантов"):
        return {"csrf_token": self.csrf(), "title": title, "direction": direction, "level": "regional",
                "event_type": "festival", "organizer": "Тестовый центр", "event_date": "2026-12-10",
                "application_deadline": "2026-11-20", "format": "hybrid", "description": "Описание события",
                "external_url": "https://example.org/event", "active": "1"}

    def test_login_and_dashboard(self):
        response = self.login()
        self.assertEqual(response.status_code, 200)
        self.assertIn("Панель достижений".encode(), response.data)
        self.assertIn("admin".encode(), response.data)

    def test_public_demo_only_exposes_demo_records(self):
        self.app.config["PUBLIC_DEMO_MODE"] = True
        with self.app.app_context():
            db = self.app.get_db()
            db.execute("UPDATE achievements SET child_name='Закрытая запись',is_demo=0 WHERE id=1")
            db.commit()
        landing = self.client.get("/")
        self.assertEqual(landing.status_code, 200)
        self.assertIn("Открыть демоверсию".encode(), landing.data)
        response = self.client.get("/dashboard")
        self.assertEqual(response.status_code, 200)
        self.assertIn("Публичный демо-режим".encode(), response.data)
        self.assertNotIn("Закрытая запись".encode(), response.data)
        for path in ("/achievements", "/children", "/mentors"):
            with self.subTest(path=path):
                public_page = self.client.get(path)
                self.assertEqual(public_page.status_code, 200)
                self.assertNotIn("Закрытая запись".encode(), public_page.data)
        self.assertEqual(self.client.get("/achievements/1").status_code, 404)
        self.assertEqual(self.client.get("/opportunities").status_code, 302)
        self.assertEqual(self.client.get("/exports/children.xlsx").status_code, 302)
        self.login()
        response = self.client.get("/achievements/1")
        self.assertEqual(response.status_code, 200)
        self.assertIn("Закрытая запись".encode(), response.data)

    def test_scoring_creation_and_aggregation(self):
        self.login()
        response = self.client.post("/achievements/new", data=self.achievement_data(), follow_redirects=True)
        self.assertEqual(response.status_code, 200)
        self.assertIn("Тестов Тимур".encode(), response.data)
        self.assertEqual(score("all_russian", "first"), 8)
        with self.app.app_context():
            children = registry_rows(self.app.get_db(), "child")
            person = next(row for row in children if row["name"] == "Тестов Тимур Павлович")
            self.assertEqual(person["total_score"], 8)
            row = self.app.get_db().execute("SELECT direction FROM achievements WHERE child_name=?", (person["name"],)).fetchone()
            self.assertEqual(row["direction"], "research_project")

    def test_permissions_and_csrf(self):
        self.add_user("viewer", "viewer")
        self.login("viewer", "password1")
        self.assertEqual(self.client.get("/achievements/new").status_code, 403)
        self.assertEqual(self.client.get("/admin/users").status_code, 403)
        self.assertEqual(self.client.get("/children").status_code, 200)
        self.assertEqual(self.client.post("/logout", data={}).status_code, 400)

    def test_ai_direction_falls_back_without_key(self):
        self.login()
        response = self.client.get("/api/suggest-direction?contest=Олимпиада по физике")
        self.assertEqual(response.status_code, 200)
        self.assertEqual(response.get_json(), {"direction": "olympiad", "source": "rules"})
        response = self.client.post("/api/achievements/1/generate-social", headers={"X-CSRF-Token": self.csrf()})
        self.assertEqual(response.status_code, 502)

    def test_ai_direction_and_social_generation(self):
        class FakeResponse:
            def __init__(self, text):
                self.data = ('{"choices":[{"message":{"content":' + __import__("json").dumps(text, ensure_ascii=False) + '}}]}').encode()

            def __enter__(self):
                return self

            def __exit__(self, *_args):
                return False

            def read(self):
                return self.data

        self.app.config["DEEPSEEK_API_KEY"] = "test-key"
        self.login()
        with patch("app.urllib.request.urlopen", side_effect=[FakeResponse("creative"), FakeResponse("Поздравляем с прекрасным результатом!")]):
            direction = self.client.get("/api/suggest-direction?contest=Фестиваль юных талантов")
            self.assertEqual(direction.get_json(), {"direction": "creative", "source": "ai"})
            generated = self.client.post("/api/achievements/1/generate-social", headers={"X-CSRF-Token": self.csrf()})
            self.assertEqual(generated.status_code, 200)
            self.assertIn("прекрасным результатом", generated.get_json()["text"])
        with self.app.app_context():
            stored = self.app.get_db().execute("SELECT generated_social FROM achievements WHERE id=1").fetchone()[0]
            self.assertIn("прекрасным результатом", stored)

    def test_exports(self):
        self.login()
        response = self.client.get("/exports/children.xlsx")
        self.assertEqual(response.status_code, 200)
        self.assertTrue(response.data.startswith(b"PK"))
        workbook = load_workbook(filename=__import__("io").BytesIO(response.data))
        self.assertEqual(workbook.active["A1"].value, "ФИО")
        doc = self.client.get("/achievements/1/social.docx")
        self.assertEqual(doc.status_code, 200)
        self.assertTrue(doc.data.startswith(b"PK"))

    def test_editor_can_crud_but_not_admin(self):
        self.add_user("editor", "editor")
        self.login("editor", "password1")
        response = self.client.post("/achievements/new", data=self.achievement_data(), follow_redirects=False)
        self.assertEqual(response.status_code, 302)
        self.assertEqual(self.client.get("/admin/audit").status_code, 403)

    def test_authenticated_pages_render(self):
        self.login()
        for path in ("/achievements", "/achievements/new", "/children", "/mentors", "/opportunities",
                     "/opportunities/new", "/admin/child-accounts", "/admin/users", "/admin/audit"):
            with self.subTest(path=path):
                self.assertEqual(self.client.get(path).status_code, 200)

    def test_child_login_and_cabinet_data(self):
        self.client.get("/login?profile=child")
        response = self.client.post("/login?profile=child", data={"profile": "child", "username": "student", "password": "student123",
                                    "csrf_token": self.csrf()}, follow_redirects=True)
        self.assertEqual(response.status_code, 200)
        self.assertIn("Мария".encode(), response.data)
        self.assertIn("10".encode(), response.data)
        self.assertIn("Созвездие талантов".encode(), response.data)
        self.assertEqual(self.client.get("/children").status_code, 302)
        with self.client.session_transaction() as session:
            self.assertIn("child_id", session)
            self.assertNotIn("user_id", session)
        self.client.post("/cabinet/logout", data={"csrf_token": self.csrf()})
        legacy = self.client.get("/cabinet/login")
        self.assertEqual(legacy.status_code, 302)
        self.assertIn("/login?profile=child", legacy.headers["Location"])

    def test_child_invitation_response_and_csrf(self):
        self.client.get("/login?profile=child")
        self.client.post("/login?profile=child", data={"profile": "child", "username": "student", "password": "student123",
                                                 "csrf_token": self.csrf()})
        cabinet = self.client.get("/cabinet")
        self.assertIn("Экологические проекты будущего".encode(), cabinet.data)
        with self.app.app_context():
            invitation_id = self.app.get_db().execute(
                "SELECT id FROM invitations WHERE child_account_id=(SELECT id FROM child_accounts WHERE username='student')"
            ).fetchone()[0]
        self.assertEqual(self.client.post(f"/cabinet/invitations/{invitation_id}/respond",
                                          data={"status": "interested"}).status_code, 400)
        response = self.client.post(f"/cabinet/invitations/{invitation_id}/respond",
                                    data={"status": "interested", "csrf_token": self.csrf()}, follow_redirects=True)
        self.assertEqual(response.status_code, 200)
        self.assertIn("Интересно".encode(), response.data)
        with self.app.app_context():
            row = self.app.get_db().execute("SELECT status,responded_at FROM invitations WHERE id=?", (invitation_id,)).fetchone()
            self.assertEqual(row["status"], "interested")
            self.assertIsNotNone(row["responded_at"])
            actions = {row[0] for row in self.app.get_db().execute(
                "SELECT action FROM audit_log WHERE entity_type IN ('child_account','invitation')").fetchall()}
            self.assertIn("child_login", actions)
            self.assertIn("child_invitation_response", actions)

    def test_staff_permissions_for_child_accounts_and_opportunities(self):
        self.add_user("viewer2", "viewer")
        self.login("viewer2", "password1")
        self.assertEqual(self.client.get("/opportunities").status_code, 200)
        self.assertEqual(self.client.get("/opportunities/new").status_code, 403)
        self.assertEqual(self.client.get("/admin/child-accounts").status_code, 403)
        self.add_user("editor2", "editor")
        self.login("editor2", "password1")
        response = self.client.post("/opportunities/new", data=self.opportunity_data(), follow_redirects=False)
        self.assertEqual(response.status_code, 302)
        self.assertEqual(self.client.get("/admin/child-accounts").status_code, 403)
        self.login()
        response = self.client.post("/admin/child-accounts", data={"csrf_token": self.csrf(), "username": "maria2",
                                    "password": "temporary9", "child_name": "Иванова Мария Сергеевна",
                                    "school": "Школа № 1"}, follow_redirects=True)
        self.assertEqual(response.status_code, 200)
        self.assertIn("maria2".encode(), response.data)
        self.assertNotIn("temporary9".encode(), response.data)

    def test_staff_and_child_usernames_cannot_conflict(self):
        self.login()
        response = self.client.post("/admin/users", data={"csrf_token": self.csrf(), "username": "student",
                                    "password": "password1", "role": "viewer"}, follow_redirects=True)
        self.assertIn("используется ребёнком".encode(), response.data)
        response = self.client.post("/admin/child-accounts", data={"csrf_token": self.csrf(), "username": "admin",
                                    "password": "password99", "child_name": "Иванов Иван Иванович",
                                    "school": "Школа № 1"}, follow_redirects=True)
        self.assertIn("используется сотрудником".encode(), response.data)

    def test_invitations_match_exact_identity_and_direction(self):
        self.login()
        self.client.post("/admin/child-accounts", data={"csrf_token": self.csrf(), "username": "exactchild",
                         "password": "password99", "child_name": "Иванова Мария Сергеевна", "school": "Школа № 1"})
        self.client.post("/admin/child-accounts", data={"csrf_token": self.csrf(), "username": "wrongschool",
                         "password": "password99", "child_name": "Иванова Мария Сергеевна", "school": "Другая школа"})
        self.client.post("/opportunities/new", data=self.opportunity_data("creative", "Творческий старт"))
        with self.app.app_context():
            db = self.app.get_db()
            invited = {row[0] for row in db.execute("""SELECT c.username FROM invitations i
                       JOIN child_accounts c ON c.id=i.child_account_id JOIN opportunities o ON o.id=i.opportunity_id
                       WHERE o.title='Творческий старт'""").fetchall()}
            self.assertIn("student", invited)
            self.assertIn("exactchild", invited)
            self.assertNotIn("wrongschool", invited)
            before = db.execute("SELECT COUNT(*) FROM invitations").fetchone()[0]
        opportunity_id = None
        with self.app.app_context():
            opportunity_id = self.app.get_db().execute("SELECT id FROM opportunities WHERE title='Творческий старт'").fetchone()[0]
        edit_data = self.opportunity_data("creative", "Творческий старт")
        self.client.post(f"/opportunities/{opportunity_id}/edit", data=edit_data)
        with self.app.app_context():
            self.assertEqual(self.app.get_db().execute("SELECT COUNT(*) FROM invitations").fetchone()[0], before)

    def test_new_achievement_backfills_matching_invitation(self):
        self.login()
        child_name = "Новый Ребёнок Сергеевич"
        school = "Школа новых достижений"
        self.client.post("/admin/child-accounts", data={"csrf_token": self.csrf(), "username": "newchild",
                         "password": "password99", "child_name": child_name, "school": school})
        self.client.post("/opportunities/new", data=self.opportunity_data("creative", "Будущие таланты"))
        with self.app.app_context():
            db = self.app.get_db()
            count = db.execute("""SELECT COUNT(*) FROM invitations i JOIN child_accounts c ON c.id=i.child_account_id
                                JOIN opportunities o ON o.id=i.opportunity_id
                                WHERE c.username='newchild' AND o.title='Будущие таланты'""").fetchone()[0]
            self.assertEqual(count, 0)
        data = self.achievement_data()
        data.update({"csrf_token": self.csrf(), "child_name": child_name, "school": school,
                     "contest_name": "Фестиваль творчества", "direction": "creative"})
        self.client.post("/achievements/new", data=data)
        with self.app.app_context():
            db = self.app.get_db()
            count = db.execute("""SELECT COUNT(*) FROM invitations i JOIN child_accounts c ON c.id=i.child_account_id
                                JOIN opportunities o ON o.id=i.opportunity_id
                                WHERE c.username='newchild' AND o.title='Будущие таланты'""").fetchone()[0]
            self.assertEqual(count, 1)


if __name__ == "__main__":
    unittest.main(verbosity=2)
