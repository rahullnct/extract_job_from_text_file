from django.shortcuts import render

# Create your views here.
import hashlib
import re

from datetime import datetime
from io import BytesIO
from pathlib import Path

from django.http import HttpResponse
from django.shortcuts import render
from django.utils import timezone
from django.utils.text import slugify

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

from .forms import JobTextUploadForm


# ============================================================
# FINAL JOB EXCEL COLUMNS
# ============================================================

JOB_COLUMNS = [
    "job_id",
    "internal_job_id",
    "job_title",
    "company_name",
    "company_website",
    "job_description",
    "skills_required",
    "experience_required",
    "experience_min_years",
    "experience_max_years",
    "salary_min",
    "salary_max",
    "salary_currency",
    "location",
    "city",
    "state",
    "country",
    "remote_type",
    "employment_type",
    "job_category",
    "posted_date",
    "expiry_date",
    "application_deadline",
    "apply_url",
    "updated_at",
    "fetched_at",
    "is_fake",
    "is_active",
    "source_url",
    "source_type",
    "source_platform",
    "score",
]


# ============================================================
# TECHNOLOGIES / SKILLS
# ============================================================

TECHNOLOGY_PATTERNS = [
    (r"\bJava\b", "Java"),
    (r"\bSpring Boot\b", "Spring Boot"),
    (r"\bSpring\b", "Spring"),
    (r"\bMaven\b", "Maven"),

    (r"\bReact(?:JS|\.js)?\b", "React"),
    (r"\bNext\s*\.?\s*JS\b", "Next.js"),
    (r"\bTypeScript\b", "TypeScript"),
    (r"\bJavaScript\b", "JavaScript"),

    (r"\bAngular\b", "Angular"),
    (r"\bVue(?:JS|\.js)?\b", "Vue.js"),

    (r"\bNode(?:JS|\.js)?\b", "Node.js"),
    (r"\bExpress(?:JS|\.js)?\b", "Express.js"),

    (r"\bPython\b", "Python"),
    (r"\bDjango\b", "Django"),
    (r"\bFlask\b", "Flask"),

    (r"\bTerraform\b", "Terraform"),
    (r"\bKubernetes\b", "Kubernetes"),
    (r"\bDocker\b", "Docker"),

    (r"\bPostgreSQL\b", "PostgreSQL"),
    (r"\bSupabase\b", "Supabase"),
    (r"\bMongoDB\b", "MongoDB"),
    (r"\bMySQL\b", "MySQL"),
    (r"\bOracle\b", "Oracle"),
    (r"\bSQL\b", "SQL"),

    (r"\bAzure\b", "Azure"),
    (r"\bAWS\b", "AWS"),
    (r"\bGCP\b", "GCP"),
    (r"\bGoogle Cloud\b", "Google Cloud"),

    (r"\bCI\s*/\s*CD\b", "CI/CD"),
    (r"\bJira\b", "Jira"),
    (r"\bGit\b", "Git"),
    (r"\bDevOps\b", "DevOps"),

    (r"\bREST(?:ful)?\s*API(?:s)?\b", "REST API"),
]

def is_header_noise_line(line):
    """
    Returns True for Shine UI text that is not job title/company name.
    """

    if not line:
        return True

    value = line.strip()
    lower = value.lower()

    exact_noise = {
        "profile",
        "actively hiring",
        "placeholder",
        "jobs for you",
        "save icon",
        "share icon",
        "save iconshare icon",
        "job details",
        "key skills",
        "recruiter details",
        "company details",
    }

    if lower in exact_noise:
        return True

    # Examples:
    # 3 weeks ago
    # 1 day ago
    # 2 months ago
    # 5 hours ago
    if re.fullmatch(
        r"\d+\s+"
        r"(minute|minutes|hour|hours|day|days|week|weeks|month|months|year|years)"
        r"\s+ago",
        lower,
    ):
        return True

    # Examples:
    # posted 3 weeks ago
    # posted 1 day ago
    if re.fullmatch(
        r"posted\s+\d+\s+"
        r"(minute|minutes|hour|hours|day|days|week|weeks|month|months|year|years)"
        r"\s+ago",
        lower,
    ):
        return True

    # Remove icon-only lines.
    if "icon" in lower and len(value.split()) <= 5:
        return True

    return False


def extract_job_header(text):
    """
    Extract job title and company name from the main Shine job header.

    Expected patterns:

    ACTIVELY HIRING
    Senior Full Stack (React + Python ) Developer Role
    Fractal
    placeholder


    OR

    Profile
    Python Developer
    Vishwa karma
    3 weeks ago
    save iconshare icon
    Job Details


    OR

    Profile
    ACTIVELY HIRING
    Python Developer with LLM - Fluper LTD
    OpenTalent
    3 weeks ago
    save iconshare icon
    Job Details
    """

    text = normalize_text(text)

    lines = [
        line.strip()
        for line in text.split("\n")
        if line.strip()
    ]

    # -----------------------------------------------------
    # Find where the actual job header ends.
    # Usually "Job Details" comes immediately after it.
    # -----------------------------------------------------

    job_details_index = None

    for index, line in enumerate(lines):

        if line.lower() == "job details":
            job_details_index = index
            break

    if job_details_index is None:
        return "", ""

    # Only examine lines before Job Details.
    before_job_details = lines[:job_details_index]

    # We normally only need the last portion of the page header.
    before_job_details = before_job_details[-20:]

    meaningful_lines = []

    for line in before_job_details:

        if is_header_noise_line(line):
            continue

        meaningful_lines.append(line)

    # -----------------------------------------------------
    # The final two meaningful lines immediately before
    # Job Details are:
    #
    # job title
    # company name
    # -----------------------------------------------------

    if len(meaningful_lines) >= 2:

        job_title = meaningful_lines[-2]
        company_name = meaningful_lines[-1]

        return job_title, company_name

    return "", ""


# ============================================================
# FILE DECODING
# ============================================================

def decode_uploaded_file(uploaded_file):

    raw_data = uploaded_file.read()

    for encoding in [
        "utf-8-sig",
        "utf-8",
        "cp1252",
        "latin-1",
    ]:

        try:
            return raw_data.decode(encoding)

        except UnicodeDecodeError:
            pass

    return raw_data.decode(
        "utf-8",
        errors="ignore",
    )


# ============================================================
# NORMALIZE TEXT
# ============================================================

def normalize_text(text):

    text = text.replace("\r\n", "\n")
    text = text.replace("\r", "\n")
    text = text.replace("\xa0", " ")

    lines = []

    for line in text.split("\n"):

        line = line.strip()

        lines.append(line)

    return "\n".join(lines)


# ============================================================
# REMOVE SIMILAR JOBS
# ============================================================

def isolate_main_job_text(text):

    markers = [
        "Similar Jobs",
        "Jobs nearby",
        "Related Jobs",
        "Trending Blogs",
        "Trending Jobs",
        "Important Links",
    ]

    lower_text = text.lower()

    positions = []

    for marker in markers:

        position = lower_text.find(
            marker.lower()
        )

        if position != -1:
            positions.append(position)

    if positions:

        text = text[
            :min(positions)
        ]

    return text.strip()


# ============================================================
# GENERIC LABEL VALUE
# ============================================================

def get_label_value(text, labels):

    if isinstance(labels, str):
        labels = [labels]

    for label in labels:

        pattern = (
            rf"(?im)^\s*"
            rf"{re.escape(label)}"
            rf"\s*:?\s*"
            rf"(?:\n\s*)?"
            rf"([^\n]+)"
        )

        match = re.search(
            pattern,
            text,
        )

        if match:

            value = (
                match.group(1)
                .strip()
            )

            if value:
                return value

    return ""


# ============================================================
# SECTION EXTRACTOR
# ============================================================

def extract_section(
    text,
    start_heading,
    end_headings,
):

    lower_text = text.lower()

    start = lower_text.find(
        start_heading.lower()
    )

    if start == -1:
        return ""

    start += len(start_heading)

    section = text[start:]

    lower_section = section.lower()

    end_positions = []

    for heading in end_headings:

        position = lower_section.find(
            heading.lower()
        )

        if position != -1:
            end_positions.append(
                position
            )

    if end_positions:

        section = section[
            :min(end_positions)
        ]

    return section.strip()


# ============================================================
# JOB TITLE
# ============================================================

def extract_job_title(text):
    job_title, company_name = extract_job_header(text)

    if job_title:
        return job_title

    # Fallback for pages containing:
    # Designation: Python Developer

    designation = get_label_value(
        text,
        "Designation",
    )

    return designation


def extract_company_name(text):
    job_title, company_name = extract_job_header(text)

    if company_name:
        return company_name

    # Fallback for pages containing:
    # Company Name: ABC Pvt Ltd

    company_name = get_label_value(
        text,
        "Company Name",
    )

    return company_name

# ============================================================
# COMPANY NAME
# ============================================================

def extract_company_name(text):

    company_name = get_label_value(
        text,
        "Company Name",
    )

    if company_name:
        return company_name

    match = re.search(
        r"""
        (?i)
        For\s+Recruiters
        \s*\n+
        [^\n]+
        \s*\n+
        ([^\n]+)
        """,
        text,
        flags=re.VERBOSE,
    )

    if match:
        return match.group(1).strip()

    return ""


# ============================================================
# JOB DESCRIPTION
# ============================================================

def extract_job_description(text):

    description = extract_section(
        text,
        "Job Description",
        [
            "Other Details",
            "Recruiter Details",
            "About Recruiter",
        ],
    )

    description = re.sub(
        r"\n{3,}",
        "\n\n",
        description,
    )

    return description.strip()


# ============================================================
# EXPERIENCE
# ============================================================

def extract_experience(text):

    patterns = [
        r"(\d+(?:\.\d+)?)\s*to\s*(\d+(?:\.\d+)?)\s*Yrs?",
        r"(\d+(?:\.\d+)?)\s*-\s*(\d+(?:\.\d+)?)\s*Yrs?",
    ]

    for pattern in patterns:

        match = re.search(
            pattern,
            text,
            flags=re.IGNORECASE,
        )

        if match:

            minimum = float(
                match.group(1)
            )

            maximum = float(
                match.group(2)
            )

            if minimum.is_integer():
                minimum = int(minimum)

            if maximum.is_integer():
                maximum = int(maximum)

            return (
                f"{minimum} to {maximum} Yrs",
                minimum,
                maximum,
            )

    # Example:
    # 5 Yrs

    match = re.search(
        r"\b(\d+(?:\.\d+)?)\s*Yrs?\b",
        text,
        flags=re.IGNORECASE,
    )

    if match:

        years = float(
            match.group(1)
        )

        if years.is_integer():
            years = int(years)

        return (
            f"{years} Yrs",
            years,
            years,
        )

    return "", "", ""


# ============================================================
# SALARY
# ============================================================

def extract_salary(text):

    # Example:
    # Rs 3.0 - 6 Lakh/Yr
    #
    # Convert to actual annual INR:
    # 3 Lakh = 300000

    pattern = (
        r"(?:Rs\.?|₹)?\s*"
        r"(\d+(?:\.\d+)?)"
        r"\s*(?:-|to)\s*"
        r"(\d+(?:\.\d+)?)"
        r"\s*(?:Lakh|LPA)"
    )

    match = re.search(
        pattern,
        text,
        flags=re.IGNORECASE,
    )

    if match:

        minimum = (
            float(match.group(1))
            * 100000
        )

        maximum = (
            float(match.group(2))
            * 100000
        )

        return (
            int(minimum),
            int(maximum),
            "INR",
        )

    return "", "", ""


# ============================================================
# LOCATION
# ============================================================

def extract_locations(text):

    # Extract between location icon and Job Description.

    match = re.search(
        r"""
        (?is)
        location\s+icon
        (.*?)
        Job\s+Description
        """,
        text,
        flags=re.VERBOSE,
    )

    cities = []

    if match:

        location_block = (
            match.group(1)
        )

        raw_locations = re.split(
            r"[,\n]+",
            location_block,
        )

        for location in raw_locations:

            location = location.strip()

            location = re.sub(
                r"^[+]\d+$",
                "",
                location,
            )

            if (
                location
                and "icon" not in location.lower()
            ):
                cities.append(location)

    # Remove duplicates.
    cities = list(
        dict.fromkeys(cities)
    )

    combined_location = ", ".join(
        cities
    )

    return combined_location, combined_location


# ============================================================
# REMOTE TYPE
# ============================================================

def detect_remote_type(text):

    lower = text.lower()

    if (
        "wfh/wfa" in lower
        or "work from home" in lower
        or "work from anywhere" in lower
        or re.search(
            r"\bremote\b",
            lower,
        )
    ):
        return "Remote"

    if "hybrid" in lower:
        return "Hybrid"

    if (
        "work from office" in lower
        or "onsite" in lower
        or "on-site" in lower
    ):
        return "Onsite"

    return ""


# ============================================================
# EMPLOYMENT TYPE
# ============================================================

def extract_employment_type(text):

    value = get_label_value(
        text,
        "Job Type",
    )

    lower = value.lower()

    mapping = [
        ("full time", "Full-time"),
        ("full-time", "Full-time"),

        ("part time", "Part-time"),
        ("part-time", "Part-time"),

        ("internship", "Internship"),

        ("contract", "Contract"),

        ("temporary", "Temporary"),

        ("freelance", "Freelance"),
    ]

    for keyword, normalized in mapping:

        if keyword in lower:
            return normalized

    return value


# ============================================================
# SKILLS
# ============================================================

def extract_skills(job_description):

    skills = []

    for pattern, skill_name in TECHNOLOGY_PATTERNS:

        if re.search(
            pattern,
            job_description,
            flags=re.IGNORECASE,
        ):

            if skill_name not in skills:
                skills.append(
                    skill_name
                )

    return "; ".join(skills)


# ============================================================
# JOB CATEGORY
# ============================================================

def detect_job_category(
    title,
    description,
):

    text = (
        f"{title} {description}"
    ).lower()

    it_keywords = [
        "developer",
        "software",
        "java",
        "python",
        "engineer",
        "frontend",
        "backend",
        "full stack",
        "devops",
        "cloud",
        "data scientist",
        "cybersecurity",
        "database",
    ]

    finance_keywords = [
        "finance",
        "accountant",
        "accounting",
        "financial analyst",
        "banking",
    ]

    marketing_keywords = [
        "marketing",
        "seo",
        "digital marketing",
        "brand manager",
    ]

    hr_keywords = [
        "human resource",
        "human resources",
        "hr manager",
        "recruiter",
        "talent acquisition",
    ]

    sales_keywords = [
        "sales",
        "business development",
        "lead generation",
    ]

    for keyword in it_keywords:

        if keyword in text:
            return "IT"

    for keyword in finance_keywords:

        if keyword in text:
            return "Finance"

    for keyword in marketing_keywords:

        if keyword in text:
            return "Marketing"

    for keyword in hr_keywords:

        if keyword in text:
            return "Human Resources"

    for keyword in sales_keywords:

        if keyword in text:
            return "Sales"

    return "Other"


# ============================================================
# POSTED DATE
# ============================================================

def extract_posted_date(text):

    match = re.search(
        r"Date:\s*(\d{1,2}/\d{1,2}/\d{4})",
        text,
        flags=re.IGNORECASE,
    )

    if not match:
        return ""

    raw_date = match.group(1)

    try:

        parsed = datetime.strptime(
            raw_date,
            "%m/%d/%Y",
        )

        return parsed.strftime(
            "%Y-%m-%d"
        )

    except ValueError:

        return raw_date


# ============================================================
# WEBSITE
# ============================================================

def extract_company_website(text):

    return get_label_value(
        text,
        [
            "Company Website",
            "Official Website",
            "Website",
        ],
    )


# ============================================================
# SOURCE PLATFORM
# ============================================================

def detect_source_platform(
    text,
    source_url,
):

    combined = (
        f"{text} {source_url}"
    ).lower()

    if (
        "shine.com" in combined
        or "shine logo" in combined
    ):
        return "Shine"

    if "linkedin.com" in combined:
        return "LinkedIn"

    if "indeed.com" in combined:
        return "Indeed"

    if "naukri.com" in combined:
        return "Naukri"

    if "greenhouse" in combined:
        return "Greenhouse"

    if "lever.co" in combined:
        return "Lever"

    return "Other"


# ============================================================
# SOURCE TYPE
# ============================================================

def determine_source_type(
    source_platform,
):

    job_portals = {
        "Shine",
        "LinkedIn",
        "Indeed",
        "Naukri",
    }

    ats_sources = {
        "Greenhouse",
        "Lever",
    }

    if source_platform in job_portals:
        return "job_portal"

    if source_platform in ats_sources:
        return "ats_api"

    return "career_page"


# ============================================================
# INITIAL SCORE
# ============================================================

def get_source_score(source_type):

    scores = {
        "ats_api": 60,
        "career_page": 40,
        "job_portal": 30,
    }

    return scores.get(
        source_type,
        0,
    )


# ============================================================
# STABLE JOB IDENTIFIER
# ============================================================

def generate_internal_job_id(
    main_text,
    source_url,
):

    if source_url:
        identity_source = source_url
    else:
        identity_source = main_text

    digest = hashlib.sha256(
        identity_source.encode(
            "utf-8"
        )
    ).hexdigest()

    return digest[:32]


# ============================================================
# MAIN EXTRACTION
# ============================================================

def extract_job_data(
    raw_text,
    source_url,
):

    text = normalize_text(
        raw_text
    )

    main_text = isolate_main_job_text(
        text
    )

    job_title = extract_job_title(
        main_text
    )

    company_name = extract_company_name(
        main_text
    )

    job_description = (
        extract_job_description(
            main_text
        )
    )

    (
        experience_required,
        experience_min,
        experience_max,
    ) = extract_experience(
        main_text
    )

    (
        salary_min,
        salary_max,
        salary_currency,
    ) = extract_salary(
        main_text
    )

    (
        location,
        city,
    ) = extract_locations(
        main_text
    )

    skills = extract_skills(
        job_description
    )

    remote_type = detect_remote_type(
        main_text
    )

    employment_type = (
        extract_employment_type(
            main_text
        )
    )

    category = detect_job_category(
        job_title,
        job_description,
    )

    posted_date = extract_posted_date(
        main_text
    )

    source_platform = (
        detect_source_platform(
            main_text,
            source_url,
        )
    )

    source_type = (
        determine_source_type(
            source_platform
        )
    )

    score = get_source_score(
        source_type
    )

    internal_job_id = (
        generate_internal_job_id(
            main_text,
            source_url,
        )
    )

    platform_prefix = (
        source_platform.lower()
        .replace(" ", "_")
    )

    job_id = (
        f"{platform_prefix}_"
        f"{internal_job_id[:16]}"
    )

    now = timezone.now().isoformat()

    job_data = {

        "job_id":
            job_id,

        "internal_job_id":
            internal_job_id,

        "job_title":
            job_title,

        "company_name":
            company_name,

        "company_website":
            extract_company_website(
                main_text
            ),

        "job_description":
            job_description,

        "skills_required":
            skills,

        "experience_required":
            experience_required,

        "experience_min_years":
            experience_min,

        "experience_max_years":
            experience_max,

        "salary_min":
            salary_min,

        "salary_max":
            salary_max,

        "salary_currency":
            salary_currency,

        "location":
            location,

        "city":
            city,

        # Not safely available from the sample.
        "state":
            "",

        # Can later be enriched.
        "country":
            "",

        "remote_type":
            remote_type,

        "employment_type":
            employment_type,

        "job_category":
            category,

        "posted_date":
            posted_date,

        "expiry_date":
            "",

        "application_deadline":
            "",

        "apply_url":
            get_label_value(
                main_text,
                [
                    "Apply URL",
                    "Application URL",
                ],
            ),

        "updated_at":
            now,

        "fetched_at":
            now,

        # Initial value.
        # Your Job Filtering AI can modify this later.
        "is_fake":
            False,

        "is_active":
            True,

        "source_url":
            source_url,

        "source_type":
            source_type,

        "source_platform":
            source_platform,

        "score":
            score,
    }

    return job_data


# ============================================================
# CREATE EXCEL
# ============================================================

def create_excel(job_data):

    workbook = Workbook()

    worksheet = workbook.active

    worksheet.title = "Job Data"

    header_fill = PatternFill(
        fill_type="solid",
        fgColor="1E3A8A",
    )

    header_font = Font(
        bold=True,
        color="FFFFFF",
    )

    # Header row
    for column_number, column_name in enumerate(
        JOB_COLUMNS,
        start=1,
    ):

        cell = worksheet.cell(
            row=1,
            column=column_number,
            value=column_name,
        )

        cell.fill = header_fill
        cell.font = header_font

        cell.alignment = Alignment(
            horizontal="center",
            vertical="center",
            wrap_text=True,
        )

    # Job row
    for column_number, column_name in enumerate(
        JOB_COLUMNS,
        start=1,
    ):

        cell = worksheet.cell(
            row=2,
            column=column_number,
            value=job_data.get(
                column_name,
                "",
            ),
        )

        cell.alignment = Alignment(
            vertical="top",
            wrap_text=True,
        )

    worksheet.freeze_panes = "A2"

    last_column = get_column_letter(
        len(JOB_COLUMNS)
    )

    worksheet.auto_filter.ref = (
        f"A1:{last_column}2"
    )

    large_columns = {
        "job_description",
        "skills_required",
        "location",
        "source_url",
        "apply_url",
    }

    for column_number, column_name in enumerate(
        JOB_COLUMNS,
        start=1,
    ):

        letter = get_column_letter(
            column_number
        )

        if column_name in large_columns:
            width = 45
        else:
            width = 20

        worksheet.column_dimensions[
            letter
        ].width = width

    worksheet.row_dimensions[1].height = 35
    worksheet.row_dimensions[2].height = 120

    output = BytesIO()

    workbook.save(output)

    output.seek(0)

    return output


# ============================================================
# DJANGO VIEW
# ============================================================

def upload_job_text(request):

    if request.method == "POST":

        form = JobTextUploadForm(
            request.POST,
            request.FILES,
        )

        if form.is_valid():

            uploaded_file = (
                form.cleaned_data[
                    "text_file"
                ]
            )

            source_url = (
                form.cleaned_data.get(
                    "source_url",
                    "",
                )
            )

            raw_text = (
                decode_uploaded_file(
                    uploaded_file
                )
            )

            job_data = extract_job_data(
                raw_text,
                source_url,
            )

            excel_file = create_excel(
                job_data
            )

            job_title = (
                job_data.get(
                    "job_title"
                )
                or Path(
                    uploaded_file.name
                ).stem
            )

            safe_name = (
                slugify(job_title)
                or "job"
            )

            filename = (
                f"{safe_name}_"
                f"job_data.xlsx"
            )

            response = HttpResponse(
                excel_file.getvalue(),
                content_type=(
                    "application/vnd."
                    "openxmlformats-officedocument."
                    "spreadsheetml.sheet"
                ),
            )

            response[
                "Content-Disposition"
            ] = (
                f'attachment; '
                f'filename="{filename}"'
            )

            return response

    else:

        form = JobTextUploadForm()

    return render(
        request,
        "extractor/upload.html",
        {
            "form": form,
        },
    )
