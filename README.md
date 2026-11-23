# Career Intelligence — Streamlit prototype

## Files
- `app.py`: Streamlit application.
- `career_role_skills_database_enhanced.xlsx`: career-role and skill database.
- `requirements.txt`: Python dependencies.

## Setup (Windows PowerShell)

1. Put `app.py`, `requirements.txt`, and `career_role_skills_database_enhanced.xlsx` in the same folder.
2. Open PowerShell in that folder.
3. Create and activate a virtual environment (optional but recommended):

```powershell
py -m venv .venv
.venv\Scripts\Activate.ps1
```

4. Install dependencies:

```powershell
py -m pip install -r requirements.txt
```

5. Start the application:

```powershell
py -m streamlit run app.py
```

Streamlit will open the application in your browser, usually at `http://localhost:8501`.

## Included workflow
1. Upload a PDF or DOCX CV.
2. Extract keyword-detected skills.
3. Review and confirm skills.
4. Generate career-role recommendations.
5. Explore career-area job-record distributions.
6. Generate a practical development plan from the selected role's skill gaps.

## Notes
- Scanned/image-only PDFs are not supported (OCR is not included).
- Skill detection is keyword-based and requires user confirmation.
- Skill-overlap scores are not hiring probabilities or validated employability scores.
- Role-level skills are aggregated from job descriptions and may not represent mandatory requirements for every vacancy.
- The application processes uploaded CV text in the current Streamlit session; it does not call an external AI API.
