from __future__ import annotations

from functools import lru_cache
from io import BytesIO
from pathlib import Path
from re import sub
from xml.sax.saxutils import escape

import pandas as pd
from fastapi import HTTPException, Request
from fastapi.responses import HTMLResponse, Response
from fastapi.templating import Jinja2Templates
from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import mm
from reportlab.platypus import Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle

base_dir = Path(__file__).resolve().parent
templates = Jinja2Templates(directory=str(base_dir / "templates"))
DATA_FILE = base_dir.parent / "school_data.xlsx"


@lru_cache(maxsize=1)
def load_school_data():
    if not DATA_FILE.exists():
        raise FileNotFoundError(f"Excel workbook not found: {DATA_FILE}")

    students = pd.read_excel(DATA_FILE, sheet_name="Students")
    teachers = pd.read_excel(DATA_FILE, sheet_name="Teachers")
    exams = pd.read_excel(DATA_FILE, sheet_name="Exams")

    for df in (students, teachers, exams):
        df.columns = [str(column).strip() for column in df.columns]

    return students, teachers, exams


students, teachers, exams = load_school_data()


def _normalize_name(value: str) -> str:
    return str(value).strip().lower()


def _build_student_summary(student_name: str, accuracy: float) -> str:
    if accuracy >= 85:
        return f"{student_name} is performing strongly across core subjects and shows excellent consistency in classwork and assessments."
    if accuracy >= 70:
        return f"{student_name} is showing steady progress with good fundamentals, and should continue strengthening weaker topic areas."
    return f"{student_name} needs focused support in key academic areas to improve consistency and overall performance."


def _format_percent(value: float) -> str:
    return f"{round(value, 1):g}%"


def _build_student_detailed_summary(row, accuracy: float, gap: float) -> str:
    student_id = int(row["student_id"])
    student_exams = exams[exams["student_id"] == student_id].copy()
    name = str(row["name"]).strip()
    class_name = str(row["class"]).strip()
    attendance = float(row["attendance"])
    extracurricular = str(row["extracurriculars"]).strip()

    if student_exams.empty:
        exam_detail = "No exam records are available, so academic accuracy cannot yet be assessed."
    else:
        total_obtained = student_exams["marks_obtained"].sum()
        total_available = student_exams["marks_total"].sum()
        subject_details = []
        for subject, subject_rows in student_exams.groupby("subject"):
            subject_total = subject_rows["marks_total"].sum()
            subject_obtained = subject_rows["marks_obtained"].sum()
            subject_accuracy = (subject_obtained / subject_total) * 100 if subject_total else 0
            subject_details.append(f"{str(subject).strip()} {_format_percent(subject_accuracy)}")
        exam_detail = (
            f"Across {len(student_exams)} exam record(s), {int(total_obtained)} of {int(total_available)} marks were earned. "
            f"Subject results are: {', '.join(subject_details)}."
        )

    support_message = (
        "Priority support is recommended because accuracy is below 70%; review the lowest subject results and provide targeted practice."
        if accuracy < 70
        else "Continue regular revision and use the lower subject results as the next focus area."
    )
    return (
        f"{name} is enrolled in {class_name} with {attendance:g}% attendance and {extracurricular} as the recorded extracurricular activity. "
        f"Exam accuracy is {_format_percent(accuracy)}, leaving a {_format_percent(gap)} gap to the 100% target. "
        f"{exam_detail} {support_message}"
    )


def _teacher_attendance(row):
    for column in ("attendance", "attendance_%", "attendance_percent"):
        if column in teachers.columns and pd.notna(row[column]):
            return float(row[column])
    return None


def _build_teacher_detailed_summary(row, accuracy: float, subject: str) -> str:
    teacher_name = str(row["name"]).strip()
    subject_name = str(subject).strip()
    subject_exams = exams[exams["subject"].astype(str).str.strip().str.lower() == subject_name.lower()]
    total_obtained = subject_exams["marks_obtained"].sum() if not subject_exams.empty else 0
    total_available = subject_exams["marks_total"].sum() if not subject_exams.empty else 0
    sessions = row.get("sessions")
    session_detail = f" The workbook records {int(sessions)} teaching sessions." if pd.notna(sessions) else ""
    attendance = _teacher_attendance(row)
    attendance_detail = f" Attendance is {_format_percent(attendance)}." if attendance is not None else ""
    focus = "Subject results need focused review and intervention planning." if accuracy < 70 else "Subject results indicate a solid foundation; continued assessment review can sustain progress."
    return (
        f"{teacher_name} teaches {subject_name}. Across {len(subject_exams)} recorded {subject_name} exam record(s), "
        f"students earned {int(total_obtained)} of {int(total_available)} available marks, giving an overall accuracy of {_format_percent(accuracy)}."
        f"{session_detail}{attendance_detail} {focus}"
    )


def _calculate_student_metrics(student_id: int):
    student_exams = exams[exams["student_id"] == student_id]
    if student_exams.empty:
        return 0.0, 100.0

    total_marks = student_exams["marks_total"].sum()
    obtained_marks = student_exams["marks_obtained"].sum()
    accuracy = round((obtained_marks / total_marks) * 100, 1) if total_marks else 0.0
    gap = round(max(0.0, 100.0 - accuracy), 1)
    return accuracy, gap


def _calculate_teacher_metrics(teacher_name: str):
    teacher_row = teachers[teachers["name"].str.strip().str.lower() == _normalize_name(teacher_name)]
    if teacher_row.empty:
        return 0.0, "N/A"

    subject = teacher_row.iloc[0]["subject"]
    subject_exams = exams[exams["subject"].str.strip().str.lower() == str(subject).strip().lower()]
    if subject_exams.empty:
        return 0.0, "N/A"

    total_marks = subject_exams["marks_total"].sum()
    obtained_marks = subject_exams["marks_obtained"].sum()
    accuracy = round((obtained_marks / total_marks) * 100, 1) if total_marks else 0.0
    return accuracy, subject


def _class_accuracy(class_name: str) -> float:
    class_students = students[students["class"].str.strip().str.lower() == class_name.strip().lower()]
    if class_students.empty:
        return 0.0
    class_ids = class_students["student_id"].tolist()
    class_exams = exams[exams["student_id"].isin(class_ids)]
    if class_exams.empty:
        return 0.0
    total_marks = class_exams["marks_total"].sum()
    obtained_marks = class_exams["marks_obtained"].sum()
    return round((obtained_marks / total_marks) * 100, 1) if total_marks else 0.0


def _subject_accuracy(subject: str) -> float:
    subject_exams = exams[exams["subject"].str.strip().str.lower() == subject.strip().lower()]
    if subject_exams.empty:
        return 0.0
    total_marks = subject_exams["marks_total"].sum()
    obtained_marks = subject_exams["marks_obtained"].sum()
    return round((obtained_marks / total_marks) * 100, 1) if total_marks else 0.0


def get_student_narrative(student_id):
    return pd.DataFrame(), "student prompt", "student review"


def get_teacher_narrative(teacher_id):
    return pd.DataFrame(), "teacher review"


def get_school_narrative():
    return pd.DataFrame(), "school review"


async def dashboard(request: Request, tab: str = "student") -> HTMLResponse:
    student_names = students["name"].astype(str).tolist()
    teacher_names = teachers["name"].astype(str).tolist()
    return templates.TemplateResponse(
        request,
        "dashboard.html",
        {"request": request, "students": student_names, "teachers": teacher_names, "tab": tab},
    )


async def get_students_api():
    rows = []
    for _, row in students.iterrows():
        student_id = int(row["student_id"])
        accuracy, gap = _calculate_student_metrics(student_id)
        rows.append(
            {
                "id": str(student_id),
                "name": str(row["name"]).strip(),
                "class": str(row["class"]).strip(),
                "attendance": f"{int(row['attendance'])}%",
                "accuracy": f"{accuracy}%",
                "extracurricular": str(row["extracurriculars"]).strip(),
                "gap": f"{gap}%",
                "summary": _build_student_detailed_summary(row, accuracy, gap),
            }
        )
    return rows


async def get_teachers_api():
    rows = []
    for _, row in teachers.iterrows():
        teacher_name = str(row["name"])
        accuracy, subject = _calculate_teacher_metrics(teacher_name)
        rows.append(
            {
                "id": str(row["teacher_id"]),
                "name": teacher_name.strip(),
                "subject": str(subject).strip(),
                "accuracy": f"{accuracy}%",
                "summary": _build_teacher_detailed_summary(row, accuracy, subject),
            }
        )
    return rows


async def get_student_report(name: str):
    if not name:
        raise HTTPException(status_code=400, detail="Student name is required")

    matched_student = students[students["name"].str.strip().str.lower() == _normalize_name(name)]
    if matched_student.empty:
        raise HTTPException(status_code=404, detail="Student not found")

    row = matched_student.iloc[0]
    student_id = int(row["student_id"])
    accuracy, gap = _calculate_student_metrics(student_id)
    return {
        "name": str(row["name"]).strip(),
        "class": str(row["class"]).strip(),
        "attendance": f"{int(row['attendance'])}%",
        "accuracy": f"{accuracy}%",
        "extracurricular": str(row["extracurriculars"]).strip(),
        "gap": f"{gap}%",
        "summary": _build_student_detailed_summary(row, accuracy, gap),
    }


async def get_teacher_report(name: str):
    if not name:
        raise HTTPException(status_code=400, detail="Teacher name is required")

    matched_teacher = teachers[teachers["name"].str.strip().str.lower() == _normalize_name(name)]
    if matched_teacher.empty:
        raise HTTPException(status_code=404, detail="Teacher not found")

    row = matched_teacher.iloc[0]
    teacher_name = str(row["name"])
    accuracy, subject = _calculate_teacher_metrics(teacher_name)
    report = {
        "name": teacher_name,
        "subject": str(subject).strip(),
        "accuracy": f"{accuracy}%",
        "summary": _build_teacher_detailed_summary(row, accuracy, subject),
    }
    attendance = _teacher_attendance(row)
    if attendance is not None:
        report["attendance"] = f"{_format_percent(attendance)}"
    return report


async def get_school_report():
    overall_accuracy = 0.0
    total_marks = exams["marks_total"].sum()
    obtained_marks = exams["marks_obtained"].sum()
    if total_marks:
        overall_accuracy = round((obtained_marks / total_marks) * 100, 1)

    overall_attendance = round(students["attendance"].mean(), 1)
    class_performance = []
    for class_name in sorted(students["class"].dropna().unique()):
        class_performance.append({
            "label": class_name,
            "value": _class_accuracy(class_name),
        })

    subject_performance = []
    for subject in sorted(exams["subject"].dropna().unique()):
        subject_performance.append({
            "label": subject,
            "value": _subject_accuracy(subject),
        })

    system_gaps = sum(1 for _, row in students.iterrows() if _calculate_student_metrics(int(row["student_id"]))[0] < 70)
    teacher_gaps = sum(1 for _, row in teachers.iterrows() if _calculate_teacher_metrics(str(row["name"]))[0] < 70)
    student_gaps = sum(1 for _, row in students.iterrows() if _calculate_student_metrics(int(row["student_id"]))[0] < 70)

    return {
        "overall_attendance": f"{overall_attendance}%",
        "overall_accuracy": f"{overall_accuracy}%",
        "system_gaps": system_gaps,
        "teacher_gaps": teacher_gaps,
        "student_gaps": student_gaps,
        "class_performance": class_performance,
        "subject_performance": subject_performance,
    }


def _build_pdf(report_type: str, report: dict) -> bytes:
    buffer = BytesIO()
    document = SimpleDocTemplate(
        buffer,
        pagesize=A4,
        rightMargin=18 * mm,
        leftMargin=18 * mm,
        topMargin=16 * mm,
        bottomMargin=16 * mm,
        title=f"The School Reports - {report_type.title()} Report",
    )
    styles = getSampleStyleSheet()
    title_style = ParagraphStyle(
        "ReportTitle",
        parent=styles["Title"],
        alignment=TA_CENTER,
        textColor=colors.HexColor("#3d5a48"),
        spaceAfter=6 * mm,
    )
    heading_style = ParagraphStyle(
        "ReportHeading",
        parent=styles["Heading2"],
        textColor=colors.HexColor("#3d5a48"),
        spaceBefore=4 * mm,
        spaceAfter=3 * mm,
    )
    body_style = ParagraphStyle(
        "ReportBody",
        parent=styles["BodyText"],
        fontSize=10.5,
        leading=15,
        textColor=colors.HexColor("#2f3b2f"),
        spaceAfter=3 * mm,
    )
    label_style = ParagraphStyle(
        "ReportLabel",
        parent=body_style,
        textColor=colors.HexColor("#5f6d5a"),
    )

    def paragraph(text, style=body_style):
        return Paragraph(escape(str(text)), style)

    story = [paragraph("The School Reports", title_style), paragraph(f"{report_type.title()} Report", heading_style)]
    profile = report.get("profile", [])
    if profile:
        profile_table = Table(
            [[paragraph(label, label_style), paragraph(value)] for label, value in profile],
            colWidths=[46 * mm, 120 * mm],
        )
        profile_table.setStyle(TableStyle([
            ("BACKGROUND", (0, 0), (-1, -1), colors.HexColor("#edf3ea")),
            ("BOX", (0, 0), (-1, -1), 0.5, colors.HexColor("#c7d3bd")),
            ("INNERGRID", (0, 0), (-1, -1), 0.25, colors.HexColor("#c7d3bd")),
            ("VALIGN", (0, 0), (-1, -1), "TOP"),
            ("LEFTPADDING", (0, 0), (-1, -1), 8),
            ("RIGHTPADDING", (0, 0), (-1, -1), 8),
            ("TOPPADDING", (0, 0), (-1, -1), 6),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 6),
        ]))
        story.extend([profile_table, Spacer(1, 5 * mm)])

    story.append(paragraph("Report Summary", heading_style))
    story.append(paragraph(report["summary"]))

    metrics = report.get("metrics", [])
    if metrics:
        story.append(paragraph("Key Metrics", heading_style))
        metric_table = Table(
            [[paragraph(label, label_style), paragraph(value)] for label, value in metrics],
            colWidths=[70 * mm, 96 * mm],
        )
        metric_table.setStyle(TableStyle([
            ("BACKGROUND", (0, 0), (-1, -1), colors.HexColor("#f5f6f2")),
            ("BOX", (0, 0), (-1, -1), 0.5, colors.HexColor("#c7d3bd")),
            ("INNERGRID", (0, 0), (-1, -1), 0.25, colors.HexColor("#c7d3bd")),
            ("LEFTPADDING", (0, 0), (-1, -1), 8),
            ("RIGHTPADDING", (0, 0), (-1, -1), 8),
            ("TOPPADDING", (0, 0), (-1, -1), 6),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 6),
        ]))
        story.append(metric_table)

    if report.get("sections"):
        for heading, content in report["sections"]:
            story.append(paragraph(heading, heading_style))
            story.append(paragraph(content))

    document.build(story)
    return buffer.getvalue()


def _download_report_payload(report_type: str, report: dict):
    filename_base = sub(r"[^A-Za-z0-9_-]+", "-", report.get("filename", report_type)).strip("-") or "school-report"
    pdf = _build_pdf(report_type, report)
    return Response(
        content=pdf,
        media_type="application/pdf",
        headers={"Content-Disposition": f'attachment; filename="{filename_base}.pdf"'},
    )


async def download_report(report_type: str, name: str = ""):
    if report_type == "student":
        data = await get_student_report(name)
        report = {
            "filename": f"student-report-{data['name']}",
            "profile": [("Student", data["name"]), ("Class", data["class"]), ("Extracurricular", data["extracurricular"])],
            "summary": data["summary"],
            "metrics": [("Attendance", data["attendance"]), ("Exam accuracy", data["accuracy"]), ("Gap to 100%", data["gap"])],
        }
    elif report_type == "teacher":
        data = await get_teacher_report(name)
        metrics = [("Subject accuracy", data["accuracy"])]
        if data.get("attendance"):
            metrics.append(("Attendance", data["attendance"]))
        report = {
            "filename": f"teacher-report-{data['name']}",
            "profile": [("Teacher", data["name"]), ("Subject", data["subject"])],
            "summary": data["summary"],
            "metrics": metrics,
        }
    elif report_type == "correspondent":
        data = await get_school_report()
        class_values = ", ".join(f"{item['label'].strip()}: {item['value']}%" for item in data["class_performance"])
        subject_values = ", ".join(f"{item['label'].strip()}: {item['value']}%" for item in data["subject_performance"])
        report = {
            "filename": "correspondent-report",
            "summary": f"Overall academic accuracy is {data['overall_accuracy']} and overall student attendance is {data['overall_attendance']}. {data['student_gaps']} student(s) and {data['teacher_gaps']} teacher(s) are below the 70% review threshold.",
            "metrics": [("Overall attendance", data["overall_attendance"]), ("Overall accuracy", data["overall_accuracy"]), ("Students below 70%", data["student_gaps"])],
            "sections": [("Class Performance", class_values or "No class exam data available."), ("Subject Accuracy", subject_values or "No subject exam data available.")],
        }
    else:
        raise HTTPException(status_code=400, detail="Report type must be student, teacher, or correspondent")
    return _download_report_payload(report_type, report)
