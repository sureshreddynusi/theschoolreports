import uvicorn

from school_reports import create_app

app = create_app()


if __name__ == "__main__":
    uvicorn.run("app:app", host="0.0.0.0", port=5000, reload=True)
