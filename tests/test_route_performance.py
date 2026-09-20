import pandas as pd

import school_reports.routes as routes
from school_reports import create_app


def test_student_tab_does_not_generate_reports_until_user_loads_it(monkeypatch):
    monkeypatch.setattr(
        routes,
        "students",
        pd.DataFrame(
            {
                "student_id": [101],
                "name": ["Alice"],
                "class": ["A"],
                "attendance": [90],
                "extracurriculars": ["Art"],
            }
        ),
    )
    monkeypatch.setattr(
        routes,
        "teachers",
        pd.DataFrame(
            {
                "teacher_id": [201],
                "name": ["Dr. Smith"],
                "subject": ["Math"],
                "sessions": [12],
            }
        ),
    )

    calls = {"student": 0, "teacher": 0, "school": 0}

    def fake_student(student_id):
        calls["student"] += 1
        return (
            pd.DataFrame([{"topic": "Math", "accuracy_%": 85.0}]),
            "student prompt",
            "student review",
        )

    def fake_teacher(teacher_id):
        calls["teacher"] += 1
        return (
            pd.DataFrame([{"topic": "Algebra", "accuracy_%": 80.0}]),
            "teacher review",
        )

    def fake_school():
        calls["school"] += 1
        return (
            pd.DataFrame([{"topic": "Math", "accuracy_%": 82.0}]),
            "school review",
        )

    monkeypatch.setattr(routes, "get_student_narrative", fake_student)
    monkeypatch.setattr(routes, "get_teacher_narrative", fake_teacher)
    monkeypatch.setattr(routes, "get_school_narrative", fake_school)

    app = create_app()
    with app.test_client() as client:
        response = client.get("/?tab=student")

    assert response.status_code == 200
    assert calls["student"] == 0
    assert calls["teacher"] == 0
    assert calls["school"] == 0
