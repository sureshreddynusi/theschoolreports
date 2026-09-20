# The School Reports

A FastAPI dashboard for generating student, teacher, and correspondent reports from the school data stored in `school_data.xlsx`.

## Features

- Student reports with:
  - Class, attendance, accuracy, and extracurricular information
  - Gap Attribution donut showing `100% - exam accuracy`
  - Attendance donut
  - Detailed data-grounded summary with subject-level results and improvement guidance
- Teacher reports with:
  - Subject and overall subject accuracy
  - Subject Accuracy donut
  - Optional attendance donut when teacher attendance exists in the workbook
  - Detailed summary using subject marks and recorded teaching sessions
- Correspondent reports with:
  - Overall attendance and exam accuracy
  - Count of students and teachers below the 70% review threshold
  - Class performance chart
  - Subject accuracy chart
- PDF downloads for the currently selected report:
  - Student: `student-report-<name>.pdf`
  - Teacher: `teacher-report-<name>.pdf`
  - Correspondent: `correspondent-report.pdf`

## Requirements

- Python 3.10 or newer
- Excel workbook at the project root: `school_data.xlsx`

Install the dependencies:

```powershell
python -m pip install -r requirements.txt
```

## Data Source

The application reads `school_data.xlsx` from the project root. The workbook currently contains these sheets:

### `Students`

Expected columns:

- `student_id`
- `name`
- `class`
- `attendance`
- `extracurriculars`

### `Teachers`

Expected columns:

- `teacher_id`
- `name`
- `subject`
- `sessions`

Teacher attendance is optional. If the sheet contains `attendance`, `attendance_%`, or `attendance_percent`, the teacher report displays an attendance donut and includes the value in the PDF.

### `Exams`

Expected columns:

- `student_id`
- `subject`
- `topic`
- `marks_obtained`
- `marks_total`

## Run the Application

From the project root:

```powershell
python app.py
```

Open the dashboard at:

```text
http://127.0.0.1:5000/
```

The application runs with Uvicorn and reloads automatically during development.

## Report Calculations

### Student accuracy

Student exam accuracy is calculated as:

```text
total marks obtained / total marks available * 100
```

### Student gap attribution

The student gap is calculated as:

```text
100 - student exam accuracy
```

### Class and subject performance

Class performance is the combined exam accuracy for students in that class. Subject accuracy is calculated from all exam records for that subject.

### Review thresholds

The correspondent report counts students and teachers below 70% accuracy:

- Students below 70%: student exam accuracy `< 70`
- Teachers below 70%: teacher subject accuracy `< 70`

## API Routes

| Route | Purpose |
| --- | --- |
| `GET /` | Render the dashboard |
| `GET /api/students` | Return student summaries and metrics |
| `GET /api/teachers` | Return teacher summaries and metrics |
| `GET /api/student-report?name=<name>` | Return one detailed student report |
| `GET /api/teacher-report?name=<name>` | Return one detailed teacher report |
| `GET /api/school-report` | Return correspondent metrics and chart data |
| `GET /api/download-report?report_type=student&name=<name>` | Download a student PDF |
| `GET /api/download-report?report_type=teacher&name=<name>` | Download a teacher PDF |
| `GET /api/download-report?report_type=correspondent` | Download a correspondent PDF |

## Project Structure

```text
app.py                         Application entry point
backend.py                     Existing backend support module
requirements.txt               Python dependencies
school_data.xlsx               Workbook used as the report data source
school_reports/
  __init__.py                  FastAPI application factory and route registration
  routes.py                    Workbook loading, calculations, reports, and PDFs
  templates/
    base.html                  Shared page styles and layout
    dashboard.html              Dashboard markup and browser behavior
tests/
  test_route_performance.py    Regression test for report loading behavior
```

## Testing

Run the test suite with:

```powershell
python -m pytest -q
```

The PDF endpoints can also be checked manually after starting the application:

```powershell
Invoke-WebRequest "http://127.0.0.1:5000/api/download-report?report_type=student&name=Aisha" -OutFile student-report.pdf
```

The downloaded file should be a PDF document and begin with the PDF file signature `%PDF-`.


###The screenshots from the application:
##Student Report
<img width="1920" height="959" alt="image" src="https://github.com/user-attachments/assets/72828901-1b3b-40ee-a332-9f036f6fab38" />

##Teacher Report
<img width="1896" height="959" alt="image" src="https://github.com/user-attachments/assets/121b7377-0fbe-4aef-9ff0-08bc6b287c5b" />

##Correspondent Report
<img width="1897" height="957" alt="image" src="https://github.com/user-attachments/assets/2c908517-dc1f-4e45-8e95-f7cfcf3b4db7" />




