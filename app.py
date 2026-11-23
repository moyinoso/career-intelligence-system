from __future__ import annotations

import io
import re
from pathlib import Path

import pandas as pd
import streamlit as st
import altair as alt

try:
    from pypdf import PdfReader
except ImportError:
    PdfReader = None

try:
    from docx import Document
except ImportError:
    Document = None


st.set_page_config(
    page_title="Career Intelligence",
    page_icon="🎯",
    layout="wide",
    initial_sidebar_state="expanded",
)

# -----------------------------
# Configuration and dictionaries
# -----------------------------
DATABASE_PATH = Path(__file__).parent / "career_role_skills_database_enhanced.xlsx"

SKILL_PATTERNS = {
    "Python": [r"\bpython\b"],
    "SQL": [r"\bsql\b", r"\bstructured query language\b"],
    "MySQL": [r"\bmysql\b"],
    "SQLite": [r"\bsqlite\b"],
    "Excel": [r"\bexcel\b", r"\bmicrosoft excel\b", r"\bms excel\b"],
    "Power BI": [r"\bpower\s*bi\b", r"\bpowerbi\b"],
    "Tableau": [r"\btableau\b"],
    "Data Analysis": [r"\bdata analysis\b", r"\bdata analytics\b", r"\bdataanalysis\b"],
    "Pandas": [r"\bpandas\b"],
    "NumPy": [r"\bnumpy\b"],
    "BeautifulSoup": [r"\bbeautiful\s?soup\b", r"\bbs4\b"],
    "Plotly": [r"\bplotly\b"],
    "DAX": [r"\bdax\b"],
    "Power Query": [r"\bpower query\b"],
    "Data Modelling": [r"\bdata model(?:ling|ing)\b"],
    "Statistical Modelling": [r"\bstatistical model(?:ling|ing)\b"],
    "Machine Learning": [r"\bmachine learning\b", r"\bscikit[\s-]?learn\b"],
    "Scikit-learn": [r"\bscikit[\s-]?learn\b", r"\bsklearn\b"],
    "JavaScript": [r"\bjavascript\b"],
    "HTML/CSS": [r"\bhtml\b", r"\bcss\b"],
    "React": [r"\breact(?:\.js)?\b"],
    "Node.js": [r"\bnode\.?js\b"],
    "MongoDB": [r"\bmongodb\b"],
    "Express.js": [r"\bexpress\.?js\b"],
    "Communication": [r"\bcommunication\b", r"\bwritten and verbal\b"],
    "Leadership": [r"\bleadership\b", r"\bled a team\b", r"\bteam leadership\b"],
    "Problem Solving": [r"\bproblem[\s-]solving\b"],
    "Teamwork": [r"\bteamwork\b", r"\bteam player\b", r"\bcollaborat(?:e|ion|ive)\b"],
    "Project Management": [r"\bproject management\b", r"\bmanaged projects\b"],
    "Research": [r"\bresearch\b"],
    "Customer Service": [r"\bcustomer service\b", r"\bcustomer support\b"],
    "Sales": [r"\bsales\b", r"\bbusiness development\b"],
    "Accounting": [r"\baccount(?:ing|ant|ancy)\b"],
    "Financial Analysis": [r"\bfinancial analysis\b", r"\bfinancial modelling\b", r"\bfinancial modeling\b"],
    "Digital Marketing": [r"\bdigital marketing\b", r"\bseo\b", r"\bsocial media marketing\b"],
}

ALIASES = {
    "ms excel": "excel",
    "microsoft excel": "excel",
    "powerbi": "power bi",
    "data analytics": "data analysis",
    "dataanalysis": "data analysis",
    "collaboration": "teamwork",
    "beautiful soup": "beautifulsoup",
    "bs4": "beautifulsoup",
    "scikit learn": "scikit-learn",
    "statistical modeling": "statistical modelling",
    "data modeling": "data modelling",
}

def normalise_skill(value: str) -> str:
    value = re.sub(r"\s+", " ", str(value).strip().lower())
    return ALIASES.get(value, value)

def extract_cv_text(uploaded_file) -> str:
    suffix = Path(uploaded_file.name).suffix.lower()
    raw = uploaded_file.getvalue()

    if suffix == ".pdf":
        if PdfReader is None:
            raise RuntimeError("PDF support is missing. Install pypdf.")
        reader = PdfReader(io.BytesIO(raw))
        text = "\n".join(page.extract_text() or "" for page in reader.pages)
    elif suffix == ".docx":
        if Document is None:
            raise RuntimeError("Word support is missing. Install python-docx.")
        doc = Document(io.BytesIO(raw))
        parts = [p.text for p in doc.paragraphs if p.text.strip()]
        for table in doc.tables:
            for row in table.rows:
                parts.append(" | ".join(cell.text.strip() for cell in row.cells))
        text = "\n".join(parts)
    else:
        raise ValueError("Upload a PDF or DOCX file.")

    # Repair common UTF-8 / Windows-1252 mojibake where possible.
    if any(token in text for token in ("â€“", "â€”", "â€¢", "Â·", "â€™")):
        try:
            text = text.encode("cp1252").decode("utf-8")
        except (UnicodeEncodeError, UnicodeDecodeError):
            pass
    return text

def detect_skills(text: str) -> list[str]:
    found = []
    for skill, patterns in SKILL_PATTERNS.items():
        if any(re.search(pattern, text, re.I) for pattern in patterns):
            found.append(skill)
    # Collapse aliases to canonical labels.
    return sorted(set(normalise_skill(x) for x in found))

def split_skill_tags(value) -> list[str]:
    if pd.isna(value) or not str(value).strip():
        return []
    return sorted({normalise_skill(x) for x in str(value).split(";") if x.strip()})

def title_seniority(title: str) -> str:
    if re.search(r"\b(intern|internship|graduate|entry[- ]level|junior|trainee|apprentice)\b", title, re.I):
        return "Graduate / Entry-level signal"
    if re.search(r"\b(head|director|chief|ceo|md\b|vice president|senior|sr\.?|lead|manager)\b", title, re.I):
        return "Senior / Leadership signal"
    return "Unspecified / Other"

def is_noisy_title(title: str) -> bool:
    patterns = [
        r"\bupfront payment\b", r"\bapply now\b", r"\bjob vacancy\b",
        r"\bjob opportunity\b", r"\burgent hiring\b", r"\bclick here\b",
        r"\bwalk[- ]in interview\b", r"https?://", r"\b\d{2,}\s*%",
        r"\bacademy\b", r"\btraining programme\b", r"\btraining program\b",
    ]
    return any(re.search(pattern, title, re.I) for pattern in patterns)

@st.cache_data(show_spinner=False)
def load_roles(path: str) -> pd.DataFrame:
    roles = pd.read_excel(path, sheet_name="Career Role Profiles")
    needed = {"career_area", "career_role", "job_records", "skill_tags"}
    missing = needed - set(roles.columns)
    if missing:
        raise ValueError(f"Career database is missing columns: {sorted(missing)}")
    return roles

def build_recommendations(roles: pd.DataFrame, skills: set[str], min_role_skills: int,
                          min_matches: int, include_senior: bool) -> pd.DataFrame:
    rows = []
    for _, row in roles.iterrows():
        title = str(row["career_role"]).strip()
        if not title or title.lower() == "nan" or is_noisy_title(title):
            continue
        seniority = title_seniority(title)
        if seniority == "Senior / Leadership signal" and not include_senior:
            continue

        role_skills = set(split_skill_tags(row["skill_tags"]))
        if len(role_skills) < min_role_skills:
            continue

        matched = sorted(skills & role_skills)
        gaps = sorted(role_skills - skills)
        if len(matched) < min_matches:
            continue

        score = len(matched) / len(role_skills) * 100
        if score >= 75:
            category = "Strong skill overlap"
        elif score >= 50:
            category = "Moderate skill overlap"
        else:
            category = "Development needed"

        records = pd.to_numeric(row["job_records"], errors="coerce")
        records = int(records) if pd.notna(records) else 0
        evidence = "Limited evidence" if records < 3 else ("Moderate evidence" if records < 10 else "More records available")

        rows.append({
            "Career area": row["career_area"],
            "Career role": title,
            "Seniority signal": seniority,
            "Skill overlap": round(score, 1),
            "Overlap category": category,
            "Matched skills": ", ".join(matched),
            "Potential skill gaps": ", ".join(gaps) if gaps else "No detected gaps",
            "Role skills assessed": len(role_skills),
            "Skills matched": len(matched),
            "Job records": records,
            "Evidence note": evidence,
        })

    if not rows:
        return pd.DataFrame()
    return pd.DataFrame(rows).sort_values(
        ["Skill overlap", "Skills matched", "Job records", "Career role"],
        ascending=[False, False, False, True],
    )

def development_actions(gaps: list[str]) -> list[tuple[str, str]]:
    resources = {
        "SQL": "Practise querying relational databases, joins, aggregation and window functions.",
        "Python": "Build small data-processing projects using Python and document your code.",
        "Power BI": "Create a dashboard with Power Query transformations, DAX measures and clear visuals.",
        "Excel": "Practise lookup functions, PivotTables, data cleaning and dashboard reporting.",
        "Data Analysis": "Complete an end-to-end analysis: define a question, clean data, analyse and communicate findings.",
        "Pandas": "Practise data cleaning, grouping, merging and exploratory analysis with Pandas.",
        "NumPy": "Practise arrays, vectorised operations and numerical calculations.",
        "Communication": "Write short insight summaries and practise explaining findings to a non-technical audience.",
        "Research": "Practise research questions, source evaluation, evidence synthesis and clear reporting.",
        "Teamwork": "Document collaborative project contributions and practise structured peer review.",
        "Problem Solving": "Use project case studies to show how you identified a problem and tested a solution.",
        "Leadership": "Document examples of coordination, initiative and responsibility.",
        "DAX": "Practise calculated columns, measures and filter context in Power BI.",
        "Power Query": "Practise importing, cleaning, reshaping and combining data.",
        "Data Modelling": "Learn table relationships, fact/dimension design and star schemas.",
        "Statistical Modelling": "Practise descriptive statistics, hypothesis testing and model interpretation.",
    }
    return [(skill, resources.get(skill, f"Find a practical project or course to build and demonstrate {skill}.")) for skill in gaps]

# -----------------------------
# Interface
# -----------------------------
st.title("🎯 Career Intelligence")
st.caption("AI-assisted CV analysis, career exploration and skills-gap assessment for graduates.")

if not DATABASE_PATH.exists():
    st.error(
        "Career database not found. Place career_role_skills_database_enhanced.xlsx "
        "in the same folder as this app."
    )
    st.stop()

try:
    role_data = load_roles(str(DATABASE_PATH))
except Exception as exc:
    st.error(f"Could not load the career database: {exc}")
    st.stop()

with st.sidebar:
    st.header("Assessment settings")
    min_role_skills = st.slider("Minimum skills in a role profile", 2, 8, 3)
    min_matches = st.slider("Minimum matching CV skills", 1, 5, 2)
    include_senior = st.checkbox("Include senior/leadership roles", value=False)
    st.divider()
    st.caption("Your CV is processed in this app session. This prototype does not submit CVs to an external AI service.")

tab_profile, tab_matches, tab_fit, tab_market, tab_plan = st.tabs(
    ["1. CV Profile", "2. Career Matches", "3. Job-Fit Assessment", "4. Labour Market", "5. Development Plan"]
)

if "cv_text" not in st.session_state:
    st.session_state.cv_text = ""
if "detected_skills" not in st.session_state:
    st.session_state.detected_skills = []
if "confirmed_skills" not in st.session_state:
    st.session_state.confirmed_skills = []
if "recommendations" not in st.session_state:
    st.session_state.recommendations = pd.DataFrame()

with tab_profile:
    st.subheader("Upload your CV")
    uploaded = st.file_uploader("Choose a PDF or DOCX CV", type=["pdf", "docx"])
    if uploaded is not None:
        if st.button("Analyse CV", type="primary"):
            try:
                extracted = extract_cv_text(uploaded)
                if not extracted.strip():
                    st.error("No selectable text found. Scanned CVs are not supported in this version.")
                else:
                    detected = detect_skills(extracted)
                    st.session_state.cv_text = extracted
                    st.session_state.detected_skills = detected
                    st.session_state.confirmed_skills = detected.copy()
                    st.success(f"CV analysed. {len(detected)} skills detected.")
            except Exception as exc:
                st.error(str(exc))

    if st.session_state.detected_skills:
        with st.container(border=True):
            st.markdown("#### Review and confirm detected skills")
            st.caption("Select the skills that accurately represent your experience. Remove any that you cannot demonstrate, or add relevant skills from the list.")
            all_known = sorted(set(SKILL_PATTERNS) | set(st.session_state.detected_skills))
            selected = st.multiselect(
                "Confirmed skills",
                options=all_known,
                default=[x for x in st.session_state.confirmed_skills if x in all_known],
                placeholder="Search and add skills...",
                help="These confirmed skills are used by the career matching and job-fit features.",
            )
            st.session_state.confirmed_skills = selected
            st.markdown(" ")
            c1, c2 = st.columns(2)
            c1.metric("Detected skills", len(st.session_state.detected_skills))
            c2.metric("Confirmed skills", len(selected))
            st.caption("A keyword mention is only a starting point. Confirm skills you can demonstrate with evidence.")
    else:
        st.info("Upload your CV to begin. Supported formats: PDF and DOCX.")

with tab_matches:
    st.subheader("Career-role recommendations")
    if not st.session_state.confirmed_skills:
        st.info("Analyse your CV and confirm your skills in the CV Profile tab first.")
    else:
        if st.button("Generate career recommendations", type="primary"):
            skill_set = {normalise_skill(x) for x in st.session_state.confirmed_skills}
            st.session_state.recommendations = build_recommendations(
                role_data, skill_set, min_role_skills, min_matches, include_senior
            )
        recs = st.session_state.recommendations
        if not recs.empty:
            m1, m2, m3 = st.columns(3)
            m1.metric("Roles assessed", len(role_data))
            m2.metric("Eligible roles", len(recs))
            m3.metric("Recommendations shown", min(15, len(recs)))
            st.dataframe(recs.head(15), use_container_width=True, hide_index=True)
            st.download_button(
                "Download recommendations (CSV)",
                data=recs.head(15).to_csv(index=False).encode("utf-8-sig"),
                file_name="career_recommendations.csv",
                mime="text/csv",
            )
            st.caption("Skill overlap is not a hiring probability or a validated employability score. Review actual vacancy requirements.")
        elif st.session_state.recommendations.empty:
            st.caption("Set your thresholds and generate recommendations.")

with tab_fit:
    st.subheader("Job-fit and skills-gap assessment")
    st.write("Paste a specific vacancy description to compare its detected skills with the skills you confirmed in your CV.")

    if not st.session_state.confirmed_skills:
        st.info("Analyse your CV and confirm your skills in the CV Profile tab first.")
    else:
        vacancy_text = st.text_area(
            "Paste job description", key="vacancy_description", height=240,
            placeholder="Paste the full job description, including responsibilities and requirements..."
        )
        if st.button("Assess job fit", type="primary", disabled=not vacancy_text.strip()):
            vacancy_skills = detect_skills(vacancy_text)
            confirmed = {normalise_skill(x) for x in st.session_state.confirmed_skills}
            required = {normalise_skill(x) for x in vacancy_skills}
            matched = sorted(confirmed & required)
            gaps = sorted(required - confirmed)
            coverage = (len(matched) / len(required) * 100) if required else 0.0
            education_terms = ["bachelor", "bsc", "b.eng", "beng", "hnd", "ond", "master", "msc", "phd", "diploma", "degree"]
            experience_patterns = [r"\b\d+\+?\s+years?\s+(?:of\s+)?experience\b", r"\bminimum\s+of\s+\d+\s+years?\b", r"\bentry[- ]level\b", r"\bgraduate\b", r"\binternship\b", r"\bintern\b"]
            education_mentions = [term for term in education_terms if term in vacancy_text.lower()]
            experience_mentions = []
            for pattern in experience_patterns:
                experience_mentions.extend(re.findall(pattern, vacancy_text, flags=re.I))
            st.session_state["job_fit_result"] = {
                "required": sorted(required), "matched": matched, "gaps": gaps,
                "coverage": round(coverage, 1), "education": education_mentions,
                "experience": sorted(set(x.strip() for x in experience_mentions)),
            }

        result = st.session_state.get("job_fit_result")
        if result:
            st.divider()
            st.markdown("#### Skill comparison")
            c1, c2, c3 = st.columns(3)
            c1.metric("Detected vacancy skills", len(result["required"]))
            c2.metric("Matched skills", len(result["matched"]))
            c3.metric("Skill overlap", f'{result["coverage"]}%')
            st.progress(min(result["coverage"] / 100, 1.0))
            left, right = st.columns(2)
            with left:
                st.markdown("**Skills found in your confirmed profile**")
                st.write(", ".join(result["matched"]) if result["matched"] else "No direct matches detected.")
            with right:
                st.markdown("**Potential skill gaps**")
                st.write(", ".join(result["gaps"]) if result["gaps"] else "No additional skills detected.")
            st.markdown("#### Education and experience mentions")
            e1, e2 = st.columns(2)
            e1.write("**Education terms found:** " + (", ".join(result["education"]) if result["education"] else "None detected"))
            e2.write("**Experience terms found:** " + (", ".join(result["experience"]) if result["experience"] else "None detected"))
            st.warning("This is a keyword comparison, not a hiring decision or complete eligibility assessment. Education and experience mentions are not verified against your CV.")

with tab_market:
    st.subheader("Labour-market overview")
    st.caption("Based on the career-role database assembled from collected job descriptions.")
    area_counts = role_data.groupby("career_area", dropna=False)["job_records"].sum().sort_values(ascending=True)
    a, b = st.columns([1.5, 1])
    with a:
        st.markdown("**Job-record distribution by career area**")
        chart_data = area_counts.rename_axis("Career area").reset_index(name="Job records")
        chart = (
            alt.Chart(chart_data)
            .mark_bar(cornerRadiusEnd=4)
            .encode(
                x=alt.X("Job records:Q", title="Job records", axis=alt.Axis(format=",d")),
                y=alt.Y("Career area:N", sort="-x", title=None, axis=alt.Axis(labelLimit=260, labelFontSize=12)),
                tooltip=[alt.Tooltip("Career area:N"), alt.Tooltip("Job records:Q", format=",d")],
            )
            .properties(height=330)
        )
        st.altair_chart(chart, use_container_width=True)
    with b:
        st.metric("Career-role profiles", len(role_data))
        st.metric("Job records represented", int(pd.to_numeric(role_data["job_records"], errors="coerce").fillna(0).sum()))
        st.metric("Career areas", role_data["career_area"].nunique())
    st.markdown("**Explore roles by career area**")
    areas = sorted(role_data["career_area"].dropna().unique().tolist())
    chosen_area = st.selectbox("Career area", areas)
    area_roles = role_data[role_data["career_area"] == chosen_area].copy()
    area_roles["job_records"] = pd.to_numeric(area_roles["job_records"], errors="coerce").fillna(0)
    show_cols = [c for c in ["career_role", "job_records", "skill_tags", "title_seniority_signal"] if c in area_roles.columns]
    st.dataframe(area_roles.sort_values("job_records", ascending=False)[show_cols].head(30),
                 use_container_width=True, hide_index=True)
    st.caption("Job-record counts reflect the collected dataset, not the total number of vacancies in Nigeria.")

with tab_plan:
    st.subheader("Personalised skills-development plan")
    recs = st.session_state.recommendations
    if recs.empty:
        st.info("Generate career recommendations first. Then select a target role to see its skill gaps.")
    else:
        role_options = recs["Career role"].tolist()
        selected_role = st.selectbox("Choose a target role", role_options)
        selected_row = recs[recs["Career role"] == selected_role].iloc[0]
        gaps_text = selected_row["Potential skill gaps"]
        gaps = [] if gaps_text == "No detected gaps" else [x.strip() for x in gaps_text.split(",") if x.strip()]
        st.write(f"**Target career area:** {selected_row['Career area']}")
        st.write(f"**Current skill overlap:** {selected_row['Skill overlap']}%")
        st.write(f"**Matched skills:** {selected_row['Matched skills'] or 'None detected'}")
        st.write("**Potential development areas:**")
        if gaps:
            for skill, action in development_actions(gaps):
                st.markdown(f"- **{skill}:** {action}")
        else:
            st.success("No additional skill gaps were detected by this keyword comparison. Review the actual vacancy requirements before drawing conclusions.")
        st.markdown("#### Suggested next actions")
        st.markdown(
            "1. Compare your CV against at least two current vacancies for this role.\n"
            "2. Select one or two skill gaps to work on first.\n"
            "3. Complete a practical project that demonstrates the skills.\n"
            "4. Update your CV with evidence of what you built or achieved."
        )

st.divider()
st.caption("Prototype notice: This system supports career exploration and skills development. It does not make hiring decisions, predict recruitment outcomes or guarantee employability.")
