from pathlib import Path

from fastapi import FastAPI
from fastapi.staticfiles import StaticFiles
from fastapi.testclient import TestClient

from .routes import (
    dashboard,
    download_report,
    get_school_report,
    get_student_report,
    get_students_api,
    get_teacher_report,
    get_teachers_api,
)


class SchoolReportsApp(FastAPI):
    def test_client(self):
        return TestClient(self)


def create_app() -> FastAPI:
    base_dir = Path(__file__).resolve().parent
    app = SchoolReportsApp(title="School Reports")
    static_dir = base_dir / "static"
    if static_dir.exists():
        app.mount("/static", StaticFiles(directory=str(static_dir)), name="static")
    app.add_api_route("/", dashboard, methods=["GET"])
    app.add_api_route("/api/students", get_students_api, methods=["GET"])
    app.add_api_route("/api/teachers", get_teachers_api, methods=["GET"])
    app.add_api_route("/api/student-report", get_student_report, methods=["GET"])
    app.add_api_route("/api/teacher-report", get_teacher_report, methods=["GET"])
    app.add_api_route("/api/school-report", get_school_report, methods=["GET"])
    app.add_api_route("/api/download-report", download_report, methods=["GET"])
    return app
