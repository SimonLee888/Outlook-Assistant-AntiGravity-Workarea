export type Metric = {
  label: string;
  value: string;
};

export type Project = {
  title: string;
  tag: string;
  summary: string;
  outcome: string;
  stack: string[];
  accent: string;
};

export type CodeSample = {
  id: string;
  title: string;
  file: string;
  category: string;
  summary: string;
  notes: string[];
  snippet: string;
};

export type ResumeItem = {
  period: string;
  title: string;
  description: string;
};

export type ResourceItem = {
  title: string;
  tag: string;
  summary: string;
  points: string[];
  downloadHref: string;
  repoHref?: string;
};

export type LabSignal = {
  label: string;
  value: string;
  detail: string;
};

export type LabTrack = {
  title: string;
  summary: string;
  points: string[];
};

export type LabRequirement = {
  title: string;
  detail: string;
};

export type LabWorkflowStep = {
  step: string;
  title: string;
  detail: string;
};

export type LabGuide = {
  title: string;
  audience: string;
  summary: string;
  requirements: string[];
  steps: string[];
  outputs: string[];
};

export const profile = {
  name: "Esteban Garzon",
  role: "CTO - Data, Automation, and AI Systems",
  heroTitle: "I design software that turns complex operations into elegant systems.",
  heroText:
    "I am a CTO with a background in data engineering, automation, business intelligence, and applied analytics. I build ETL pipelines, AI-enabled workflows, and internal platforms that turn operational complexity into clear, high-leverage systems.",
  resumeSummary:
    "CTO, data engineer, and automation specialist with a multidisciplinary background spanning civil engineering, data science, business intelligence, and operational systems. Experienced in building ETL pipelines, analytics layers, AI-driven automations, dashboards, and backend workflows that improve decision-making, reduce manual work, and create reliable infrastructure for growing teams.",
  linkedin: "https://www.linkedin.com/in/esteban-garzon2",
  email: "estebandre2@hotmail.es",
  phone: "+593996788376",
  resumeHref: "/downloads/resume/Esteban_Garzon_CV_Final.pdf",
};

export const metrics: Metric[] = [
  { label: "Focus", value: "Automation + Data Systems" },
  { label: "Specialty", value: "ETL, integrations, reliability" },
  { label: "Approach", value: "Operational software with premium UX" },
];

export const projects: Project[] = [
  {
    title: "Slab ETL Nexus",
    tag: "Data Pipeline",
    summary:
      "Production ETL that blends SQL Server, shop-floor Excel logs, and PostgreSQL into a trustworthy reporting layer for operations.",
    outcome:
      "Defensive deduplication, slab-level enrichment, atomic exports, and freshness checks that reduce operational friction and data drift.",
    stack: ["Python", "PostgreSQL", "SQL Server", "ETL", "Operations"],
    accent: "var(--cyan)",
  },
  {
    title: "Moraware Sync Engine",
    tag: "Sync Infrastructure",
    summary:
      "A multiprocessing sync engine that collects fabrication, CAD, and install activity and restructures it into analytics-ready records.",
    outcome:
      "Parallel processing, resilient logging, notes normalization, and vendor resolution to surface cleaner business signals from noisy source systems.",
    stack: ["Python", "PostgreSQL", "API Sync", "Concurrency", "Monitoring"],
    accent: "var(--amber)",
  },
  {
    title: "Ops Board Bridge",
    tag: "Workflow Automation",
    summary:
      "A data bridge that transforms operational status into structured task cards and visible execution layers for the team.",
    outcome:
      "Less manual coordination, clearer priorities, and a software layer that connects field activity with visual workflow management.",
    stack: ["Python", "SQLite", "Automation", "Dashboards", "Systems Design"],
    accent: "var(--rose)",
  },
];

export const capabilities = [
  "Robust ETL pipelines",
  "API and ERP integrations",
  "Multi-source data synchronization",
  "SQL analytics and data modeling",
  "Operational automation",
  "Business intelligence dashboards",
  "Machine learning and predictive analytics",
  "Interactive frontends for complex systems",
];

export const resources: ResourceItem[] = [
  {
    title: "Outlook Attachment Extractor",
    tag: "Free Utility",
    summary:
      "A Windows-based OST attachment extraction utility with GUI and CLI workflows, designed to browse Outlook data files, filter messages, and export attachments in a structured way.",
    points: [
      "Sanitized public package with source code only.",
      "No client data, extracted mail content, temporary mailboxes, or build artifacts included.",
      "Useful as a practical download for teams that need OST attachment extraction workflows.",
    ],
    downloadHref: "/downloads/outlook-attachment-extractor-source.zip",
    repoHref: "https://github.com/estebandre2/esteban-garzon-portfolio",
  },
  {
    title: "Live Data Integration Lab",
    tag: "Public Source Bundle",
    summary:
      "A curated package of sanitized integration scripts covering Moraware, RFMS, Microsoft Graph, Samsara, and Excel-based ingestion patterns.",
    points: [
      "Built from real operational automations, rewritten to remove secrets, internal IDs, private paths, and raw business data.",
      "Useful for demonstrating ETL, API authentication, batch backfills, desktop automation, and schema-shaping patterns.",
      "Safe to share as a public engineering sample without exposing the original `livedataproj` workspace.",
    ],
    downloadHref: "/downloads/live-data-integration-lab-source.zip",
    repoHref: "https://github.com/estebandre2/esteban-garzon-portfolio",
  },
];

export const resumeHighlights = [
  "Currently operating at CTO level, leading technical direction across data, automation, and AI-driven systems.",
  "Built production ETL pipelines connecting SQL Server, Excel-based shop data, PostgreSQL, APIs, and operational platforms.",
  "Designed synchronization workflows for manufacturing, CAD, installation, finance, and vendor intelligence.",
  "Experienced with Power BI, Python, SQL, machine learning, and automation-first system design.",
  "Able to bridge technical execution, strategic thinking, and multidisciplinary business contexts.",
];

export const resumeItems: ResumeItem[] = [
  {
    period: "Latest Position",
    title: "Chief Technology Officer (CTO)",
    description:
      "Leading technical strategy across automation, AI, business intelligence, and data systems while shaping solutions that improve operational visibility, efficiency, and decision-making.",
  },
  {
    period: "2024 - Present",
    title: "Data Engineer | AI-Driven Automation & Business Intelligence Specialist",
    description:
      "Designed and implemented ETL pipelines, API integrations, reporting systems, and AI-enabled automations for a U.S.-based client, connecting platforms such as Moraware, RFMS, Slabsmith, Samsara, and production machine logs.",
  },
  {
    period: "2022 - 2023",
    title: "Data Scientist, Data Analyst, and Financial Modeling Roles",
    description:
      "Worked across hydroelectric, photovoltaic, construction, and water system projects, applying data analysis, forecasting, financial modeling, and feasibility studies to support strategic decisions.",
  },
  {
    period: "Education",
    title: "Civil Engineer + Master's in Data Science and Business Intelligence",
    description:
      "Blending engineering discipline with analytics, machine learning, business intelligence, and applied data strategy to solve real-world operational problems.",
  },
];

export const codeLabSignals: LabSignal[] = [
  {
    label: "Public Samples",
    value: "13",
    detail: "Curated excerpts across ETL, APIs, telemetry, Graph, Excel automation, schema bridges, and batching flows.",
  },
  {
    label: "Integration Families",
    value: "5",
    detail: "Moraware, RFMS, Microsoft Graph, Samsara, and spreadsheet-driven shop-floor pipelines.",
  },
  {
    label: "Download Bundles",
    value: "2",
    detail: "Sanitized public utilities and source packages that are safe to share in a portfolio context.",
  },
  {
    label: "Usage Guides",
    value: "2",
    detail: "Step-by-step instructions, prerequisites, and expected outputs for the public downloads.",
  },
];

export const codeLabTracks: LabTrack[] = [
  {
    title: "Operational Ingestion",
    summary:
      "Patterns for turning shop-floor logs, ERP signals, and structured exports into cleaner analytical layers.",
    points: [
      "Spreadsheet normalization and sheet harmonization.",
      "Controlled joins between job, slab, and activity records.",
      "Incremental exports that preserve business context without leaking raw internals.",
    ],
  },
  {
    title: "API Orchestration",
    summary:
      "Session handling, credential flows, pagination, and defensive parsing across vendor APIs and cloud services.",
    points: [
      "RFMS session bootstrap and estimate hydration.",
      "Microsoft Graph client-credentials token exchange.",
      "Samsara polling windows built for historical backfills.",
    ],
  },
  {
    title: "Reliability Layers",
    summary:
      "The code emphasizes safe writes, deduplication, placeholder handling, and shape checks before records hit storage.",
    points: [
      "Completeness scoring before duplicate removal.",
      "Batch inserts and time-window chunking for long runs.",
      "Selective updates instead of blunt full-table churn where possible.",
    ],
  },
  {
    title: "Desktop + Internal Tooling",
    summary:
      "Not everything lives in a web app. Some of the strongest leverage comes from automating the awkward middle ground.",
    points: [
      "PythonNET bridges into .NET-based vendor tooling.",
      "Excel automation with `xlwings` and downstream dataframe cleanup.",
      "Utilities packaged for non-technical operators when needed.",
    ],
  },
];

export const codeLabRequirements: LabRequirement[] = [
  {
    title: "Python 3.11+",
    detail:
      "Most samples assume a modern Python runtime and common data tooling such as `requests`, `pandas`, and `psycopg2-binary`.",
  },
  {
    title: "Windows-first utilities",
    detail:
      "The Outlook tool and some spreadsheet or desktop flows are designed around Windows environments and local desktop automation.",
  },
  {
    title: "Service credentials",
    detail:
      "Public packages replace all secrets with placeholders. To run them, you must provide your own API keys, database credentials, and tenant identifiers.",
  },
  {
    title: "External systems",
    detail:
      "Some samples expect access to third-party systems such as Moraware, Microsoft Graph, RFMS, SQL Server, PostgreSQL, or Samsara.",
  },
];

export const codeLabWorkflow: LabWorkflowStep[] = [
  {
    step: "01",
    title: "Choose a bundle",
    detail:
      "Pick the public package that matches your need: Outlook extraction for email attachments or Live Data Integration Lab for ETL and API examples.",
  },
  {
    step: "02",
    title: "Read the guide first",
    detail:
      "Each bundle is sanitized and intentionally selective. Start with the README to understand the expected environment and what has been redacted.",
  },
  {
    step: "03",
    title: "Install prerequisites",
    detail:
      "Create a clean environment, install the listed packages, and replace placeholders like `***` with your own safe local configuration.",
  },
  {
    step: "04",
    title: "Run a narrow test",
    detail:
      "Begin with a single script or a reduced date range so you can validate connectivity, schema behavior, and output shape before scaling up.",
  },
  {
    step: "05",
    title: "Adapt to your stack",
    detail:
      "The samples are meant to be templates. Rename tables, adjust schemas, and route outputs to your own storage and monitoring setup.",
  },
];

export const codeLabGuides: LabGuide[] = [
  {
    title: "How to use the Live Data Integration Lab",
    audience: "For engineers exploring ETL, API sync, and data-platform patterns",
    summary:
      "This package is best used as a reference implementation. It shows structure, flow control, and integration techniques rather than pretending to be a plug-and-play app.",
    requirements: [
      "Python 3.11 or newer",
      "A virtual environment",
      "Your own PostgreSQL or SQL-compatible destination",
      "Your own vendor credentials for whichever script you want to test",
    ],
    steps: [
      "Download the zip and extract it into a clean folder.",
      "Create a virtual environment and install `requirements.txt`.",
      "Open the script you want to study first, such as the RFMS bridge or the Microsoft Graph sync.",
      "Replace the placeholder values with your own local configuration.",
      "Run a small test scope first, then expand the date range or dataset size once the shape looks correct.",
    ],
    outputs: [
      "Understanding of auth flows, polling windows, and ETL control patterns",
      "Reusable code structure for ingestion jobs",
      "A safe starting point for adapting to your own systems",
    ],
  },
  {
    title: "How to use the Outlook Attachment Extractor",
    audience: "For teams or operators who need a practical attachment-export utility",
    summary:
      "This package is intended as a usable Windows utility sample. It demonstrates both packaging discipline and a real-world local workflow.",
    requirements: [
      "Windows environment",
      "Python installed locally if running from source",
      "Access to an OST file you are authorized to inspect",
      "Enough local disk space for exported attachments",
    ],
    steps: [
      "Download the source package and extract it locally.",
      "Review the README and install the dependencies if you plan to run it from source.",
      "Launch the GUI or CLI entrypoint depending on your workflow.",
      "Point the tool to your OST file, filter by the criteria you need, and choose an export folder.",
      "Verify a small sample export first before running a broader extraction.",
    ],
    outputs: [
      "Structured attachment exports",
      "A reusable local utility for mailbox recovery or audit workflows",
      "A reference implementation for GUI-driven Python tooling",
    ],
  },
];

export const codeSamples: CodeSample[] = [
  {
    id: "dedupe",
    title: "Inventory Deduplication Logic",
    file: "slab_project/etl_slabsmith.py",
    category: "ETL Reliability",
    summary:
      "A practical deduplication pass that protects downstream merges from repeated inventory identifiers.",
    notes: [
      "Ranks rows by completeness before removing duplicates.",
      "Preserves rows without a usable key instead of discarding them.",
      "Makes the ETL safer for reporting and enrichment joins.",
    ],
    snippet: `def deduplicate_by_inventory_id(df, key_col):
    if df.empty or key_col not in df.columns:
        return df.copy(), 0

    deduped = df.copy()
    deduped[key_col] = deduped[key_col].apply(normalize_inventory_key)

    mask_has_key = deduped[key_col].notna()
    with_key = deduped[mask_has_key].copy()
    without_key = deduped[~mask_has_key].copy()

    if with_key.empty:
        return deduped, 0

    completeness_score = pd.Series(0, index=with_key.index, dtype="int64")
    for col in [col for col in with_key.columns if col != key_col]:
        if pd.api.types.is_numeric_dtype(with_key[col]):
            completeness_score += with_key[col].notna().astype("int64")
        else:
            completeness_score += (
                with_key[col].fillna("").astype(str).str.strip().ne("").astype("int64")
            )

    with_key["_dedupe_score"] = completeness_score
    with_key = with_key.sort_values([key_col, "_dedupe_score"], ascending=[True, False])
    duplicate_rows_removed = len(with_key) - with_key[key_col].nunique()
    with_key = with_key.drop_duplicates(subset=[key_col], keep="first")

    deduped = pd.concat([with_key.drop(columns=["_dedupe_score"]), without_key], ignore_index=True)
    return deduped, duplicate_rows_removed`,
  },
  {
    id: "sync",
    title: "Parallel Job Processing",
    file: "moraware_project/moraware_sync.py",
    category: "Concurrency",
    summary:
      "A multiprocessing pattern that spreads job extraction across workers and streams inserts through dedicated consumer threads.",
    notes: [
      "Separates API calls from insert throughput concerns.",
      "Uses shared counters and queues for progress reporting.",
      "Scales long-running sync jobs more gracefully.",
    ],
    snippet: `for i, chunk in enumerate(chunks):
    p = multiprocessing.Process(target=worker_process, args=(i, chunk, result_queue))
    p.start()
    processes.append(p)

while workers_done < len(chunks):
    msg = result_queue.get(timeout=5)
    msg_type = msg[0]
    worker_id = msg[1]

    if msg_type == "progress":
        w_processed = msg[2]
        prev = worker_progress.get(worker_id, 0)
        delta = w_processed - prev
        worker_progress[worker_id] = w_processed

        if worker_id in bars:
            bars[worker_id].update(delta)
    elif msg_type == "done":
        workers_done += 1
        total_processed += msg[2]
        total_inserted += msg[3]
        total_errors += msg[4]`,
  },
  {
    id: "merge",
    title: "Controlled Enrichment Layer",
    file: "slab_project/build_cad_table.py",
    category: "Data Enrichment",
    summary:
      "A controlled enrichment step that connects fabrication activity with slab metadata and operator information.",
    notes: [
      "Combines Moraware activity notes with slab lookup tables.",
      "Keeps the final output business-friendly instead of exposing full internals.",
      "Matches the same selective-access philosophy used in this portfolio.",
    ],
    snippet: `for rec in records:
    sid = rec["slab_id"]
    slab_info = slab_map.get(sid)
    operator_info = operator_map.get(sid)
    operator_shop = normalize_operator(
        operator_info["operator_shop"] if operator_info else None
    )

    if slab_info:
        data = {
            **rec,
            "inventoryid": slab_info["inventoryid"],
            "JobName": slab_info["JobName"],
            "Utilized": slab_info["Utilized"],
            "Material": slab_info["Material"],
            "operator_shop": operator_shop,
        }
    else:
        data = {
            **rec,
            "inventoryid": sid,
            "JobName": None,
            "Utilized": None,
            "Material": None,
            "operator_shop": operator_shop,
        }

    pg_cur.execute(insert_sql, data)`,
  },
  {
    id: "moraware-bridge",
    title: "Moraware Activity Bridge",
    file: "live-data-integration-lab/moraware_activity_sync.py",
    category: ".NET Interop",
    summary:
      "A PythonNET flow that connects to a .NET vendor SDK, shapes activity metadata, and persists job-level signals for reporting.",
    notes: [
      "Bridges Python and .NET rather than relying on flat exports.",
      "Converts vendor date objects into safe ISO timestamps.",
      "Extracts structured fields from noisy activity notes.",
    ],
    snippet: `def convert_to_iso_string(value):
    if value is None:
        return None
    if isinstance(value, datetime):
        return value.isoformat()
    try:
        return datetime(
            value.Year, value.Month, value.Day, value.Hour, value.Minute, value.Second
        ).isoformat()
    except AttributeError:
        return None

for activity in job_activities:
    activity_name = activity.JobActivityTypeName
    if activity_name in activity_data:
        if "assigned_to" in activity_data[activity_name]:
            assignees = []
            enumerator = activity.Assignees.GetEnumerator()
            while enumerator.MoveNext():
                assignee = enumerator.Current
                assignees.append(str(assignee.AssigneeName))
            activity_data[activity_name]["assigned_to"] = ", ".join(assignees)

        if "date" in activity_data[activity_name]:
            activity_data[activity_name]["date"] = convert_to_iso_string(activity.StartDate)`,
  },
  {
    id: "rfms-session",
    title: "RFMS Session + Estimate Hydration",
    file: "live-data-integration-lab/rfms_estimate_bridge.py",
    category: "API Ingestion",
    summary:
      "A vendor session bootstrap followed by estimate-level hydration and schema-friendly extraction for local storage.",
    notes: [
      "Uses Basic auth only for session creation, then pivots into a short-lived session token.",
      "Parses nested customer fields defensively.",
      "Computes useful derived values like total square-foot area.",
    ],
    snippet: `def get_session_token():
    auth_string = f"{STORE_QUEUE}:{API_KEY}"
    auth_base64 = base64.b64encode(auth_string.encode()).decode()
    headers = {"Authorization": f"Basic {auth_base64}"}

    response = requests.post(f"{BASE_URL}/session/begin", headers=headers, timeout=30)
    response.raise_for_status()
    return response.json().get("sessionToken")

def get_estimate_data(estimate_number, session_token):
    headers = {
        "Authorization": f"Basic {base64.b64encode(f'{STORE_QUEUE}:{session_token}'.encode()).decode()}",
        "Content-Type": "application/json",
    }

    response = requests.get(f"{BASE_URL}/estimate/{estimate_number}?locked=false", headers=headers, timeout=30)
    response.raise_for_status()
    result = response.json().get("result") or {}
    lines = result.get("lines", [])
    total_area = sum(line.get("quantity", 0) for line in lines if line.get("saleUnits", "").upper() == "SF")
    return {
        "job_number": result.get("jobNumber"),
        "line_count": len(lines),
        "total_area": total_area,
    }`,
  },
  {
    id: "teams-graph",
    title: "Microsoft Graph Timecard Sync",
    file: "live-data-integration-lab/teams_timecards_sync.py",
    category: "Cloud Automation",
    summary:
      "An OAuth client-credentials flow that exchanges app credentials for a token and pulls shift/timecard records from Microsoft Graph.",
    notes: [
      "Uses app-level auth for service-style automation.",
      "Keeps token acquisition separate from data collection.",
      "Demonstrates comfort with enterprise APIs beyond simpler webhook use cases.",
    ],
    snippet: `def get_access_token():
    token_url = f"https://login.microsoftonline.com/{TENANT_ID}/oauth2/v2.0/token"
    data = {
        "grant_type": "client_credentials",
        "client_id": CLIENT_ID,
        "client_secret": CLIENT_SECRET,
        "scope": "https://graph.microsoft.com/.default",
    }
    response = requests.post(token_url, data=data, timeout=30)
    response.raise_for_status()
    return response.json()["access_token"]

def get_time_cards(access_token):
    headers = {
        "Authorization": f"Bearer {access_token}",
        "Content-Type": "application/json",
    }
    response = requests.get(f"{GRAPH_URL}/teams/{TEAM_ID}/schedule/timeCards", headers=headers, timeout=30)
    response.raise_for_status()
    return response.json().get("value", [])`,
  },
  {
    id: "samsara-backfill",
    title: "Telemetry Backfill Windowing",
    file: "live-data-integration-lab/samsara_history_backfill.py",
    category: "Telemetry Pipelines",
    summary:
      "A time-sliced polling loop that backfills vehicle history in manageable windows rather than making one brittle bulk request.",
    notes: [
      "Chunks historical ranges into predictable intervals.",
      "Supports long-running recovery jobs and partial reruns.",
      "Pairs cleanly with batch inserts downstream.",
    ],
    snippet: `def fetch_and_store_historical_data(start_date, end_date):
    current_start = start_date
    while current_start < end_date:
        current_end = min(current_start + timedelta(minutes=30), end_date)
        start_ms = int(current_start.timestamp() * 1000)
        end_ms = int(current_end.timestamp() * 1000)

        vehicle_data = get_vehicle_data(start_ms, end_ms)
        if vehicle_data:
            save_to_postgres(vehicle_data)

        current_start = current_end`,
  },
  {
    id: "excel-normalize",
    title: "Excel Cut Log Normalization",
    file: "live-data-integration-lab/excel_cutlog_normalizer.py",
    category: "Spreadsheet Automation",
    summary:
      "A multi-sheet cleanup flow that standardizes inconsistent headers and merges shop-floor exports into one usable dataset.",
    notes: [
      "Maps variant column names into a unified schema.",
      "Keeps only the fields needed for downstream analysis.",
      "Shows practical Excel automation where APIs do not exist.",
    ],
    snippet: `COLUMN_ALIASES = {
    "Date": ["Date", "DATE"],
    "SF Utilized": ["SF Utilized", "SFT"],
    "Material Name": ["Marerial Name", "Material Name"],
    "Job Name": ["Job Name", "JOB NAME ", "JOB NAME"],
}

def normalize_columns(df, aliases):
    mapped = {}
    for canonical_name, variants in aliases.items():
        for variant in variants:
            if variant in df.columns:
                mapped[variant] = canonical_name
    return df.rename(columns=mapped)

for sheet in TARGET_SHEETS:
    df = pd.read_csv(export_path_for(sheet))
    df = normalize_columns(df, COLUMN_ALIASES)
    df = df[[col for col in COLUMN_ALIASES if col in df.columns]]
    df["Source"] = sheet
    frames.append(df)`,
  },
  {
    id: "duplicate-cleanup",
    title: "Normalized Duplicate Cleanup",
    file: "live-data-integration-lab/job_cleanup.py",
    category: "Data Quality",
    summary:
      "A cleanup pass that compares normalized job names in SQL so duplicate operational records can be removed deterministically.",
    notes: [
      "Uses trimmed, lower-cased keys instead of fragile direct matches.",
      "Keeps one survivor row per logical job.",
      "Pairs raw ingestion with controlled post-load cleanup.",
    ],
    snippet: `find_duplicates_query = """
SELECT LOWER(TRIM(job_name)) AS normalized_name, COUNT(*)
FROM job_data
GROUP BY LOWER(TRIM(job_name))
HAVING COUNT(*) > 1;
"""

delete_query = """
DELETE FROM job_data
WHERE ctid IN (
    SELECT ctid
    FROM job_data
    WHERE LOWER(TRIM(job_name)) = LOWER(TRIM(%s))
      AND ctid NOT IN (
          SELECT MIN(ctid)
          FROM job_data
          WHERE LOWER(TRIM(job_name)) = LOWER(TRIM(%s))
      )
);
"""`,
  },
  {
    id: "schema-bootstrap",
    title: "Schema Bootstrap for Operational Tables",
    file: "livedataproj/RFMSorders.py",
    category: "Database Design",
    summary:
      "A lightweight schema bootstrap pattern that creates the target table before ingestion and avoids brittle manual setup.",
    notes: [
      "Ensures the destination exists before the fetch loop starts.",
      "Useful for internal tools that must be deployable by operators.",
      "Pairs cleanly with insert-on-conflict or append-only workflows.",
    ],
    snippet: `create_table_query = """
CREATE TABLE IF NOT EXISTS purchase_orders_rfms (
    id SERIAL PRIMARY KEY,
    document_number VARCHAR(255) NOT NULL
);
"""

alter_queries = [
    "DO $$ BEGIN IF NOT EXISTS (SELECT 1 FROM information_schema.columns WHERE table_name='purchase_orders_rfms' AND column_name='document_number') THEN ALTER TABLE purchase_orders_rfms ADD COLUMN document_number VARCHAR(255) NOT NULL; END IF; END $$;"
]

cursor.execute(create_table_query)
for query in alter_queries:
    cursor.execute(query)
    connection.commit()`,
  },
  {
    id: "scheduler-orchestration",
    title: "Scheduled Pipeline Orchestration",
    file: "livedataproj/Scheduledopen2.py",
    category: "Automation Ops",
    summary:
      "A simple but effective scheduler that chains multiple scripts into one recurring operational workflow.",
    notes: [
      "Good example of pragmatic orchestration before introducing a full scheduler stack.",
      "Useful for internal recurring jobs that need fast deployment.",
      "Shows comfort with turning separate scripts into a repeatable pipeline.",
    ],
    snippet: `def ejecutar_programa():
    try:
        comando1 = ["python", "Slabsmithfinal.py"]
        subprocess.run(comando1, check=True)

        comando2 = ["python", "cutlogsconv.py"]
        subprocess.run(comando2, check=True)

        comando3 = ["python", "cutlogsinslabsmith.py"]
        subprocess.run(comando3, check=True)

        hora_actual = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        print(f"Programas ejecutados exitosamente a las {hora_actual}.")
    except subprocess.CalledProcessError as e:
        print(f"Error al ejecutar un programa: {e}")

schedule.every(15).minutes.do(ejecutar_programa)`,
  },
  {
    id: "sqlserver-bridge",
    title: "SQL Server to PostgreSQL Bridge",
    file: "livedataproj/Slabsmithfinal.py",
    category: "Cross-Database Sync",
    summary:
      "A cross-database bridge that reads from SQL Server, reshapes units, and lands the result into PostgreSQL for downstream analytics.",
    notes: [
      "Practical pattern for moving line-of-business data out of an operational database.",
      "Includes unit conversion before persistence so downstream consumers get analytics-ready fields.",
      "Demonstrates comfort with ODBC, pandas, and SQLAlchemy in one flow.",
    ],
    snippet: `query = """
SELECT
    slabs.[InventoryID],
    slabs.[Material],
    slabs.[Manufacturer],
    slabs.[Product],
    slabs.[Area$Usable],
    jobslabs.[Utilized],
    jobslabs.[Reserved],
    jobslabs.[Waste],
    jobs.[ID] AS JobName
FROM [SSSlabBrowser].[All Slabs] slabs
INNER JOIN [dbo].[JobSlabs] jobslabs ON slabs.[SlabID] = jobslabs.[SlabUID]
INNER JOIN [dbo].[JobItems] jobitems ON jobslabs.[JobItemUID] = jobitems.[JobItemID]
INNER JOIN [dbo].[Jobs] jobs ON jobitems.[JobUID] = jobs.[JobID];
"""

for column in ["Area$Usable", "Utilized", "Reserved", "Waste"]:
    data[column] = data[column] * 10.7639

data.to_sql("job_slabs_data", engine, if_exists="replace", index=False)`,
  },
  {
    id: "batch-insert",
    title: "Batch Event Insert Pattern",
    file: "livedataproj/Samsara2.py",
    category: "Write Throughput",
    summary:
      "A batching pattern that accumulates event rows and persists them in one database operation instead of writing one row at a time.",
    notes: [
      "Reduces overhead on long-running event ingestion jobs.",
      "Improves clarity by separating event parsing from persistence.",
      "Good foundation for later swapping in COPY or queue-based writes.",
    ],
    snippet: `def guardar_eventos_lote(entries):
    if not entries:
        print("No hay eventos para guardar.")
        return

    cursor.executemany(
        """
        INSERT INTO vehiculos_ubicacion (vehiculo_id, tipo_evento, timestamp, latitud, longitud, ubicacion)
        VALUES (%s, %s, %s, %s, %s, %s)
        """,
        entries
    )
    conn.commit()

for evento in eventos:
    tipo_evento = evento.get("type", "Desconocido")
    timestamp = datetime.fromisoformat(evento["time"].replace("Z", "+00:00"))
    ubicacion = evento.get("location", {})
    entries.append((vehiculo_id, tipo_evento, timestamp, latitud, longitud, direccion))`,
  },
];

export type AnalyticsSignal = {
  label: string;
  value: string;
  detail: string;
};

export type AnalyticsProject = {
  title: string;
  tag: string;
  summary: string;
  outcome: string;
  stack: string[];
  accent: string;
  href: string;
  hrefLabel: string;
};

export type AnalyticsArchiveItem = {
  title: string;
  tag: string;
  summary: string;
  format: string;
  href: string;
  hrefLabel?: string;
};

export type AnalyticsCaseStudy = {
  slug: string;
  title: string;
  tag: string;
  summary: string;
  overview: string;
  whatItShows: string[];
  focusAreas: string[];
  deliverables: string[];
};

export const analyticsSignals: AnalyticsSignal[] = [
  {
    label: "Portfolio Layer",
    value: "20",
    detail: "Visible analytics deliverables across BI, visualization, graph analysis, applied reporting, databases, optimization, and strategy briefs.",
  },
  {
    label: "Featured Cases",
    value: "6",
    detail: "The strongest pieces are separated from the broader archive so the narrative stays sharp and every link stays live.",
  },
  {
    label: "Download Types",
    value: "PDF / PBIX / TWB / PPTX / DOCX / Briefs",
    detail: "The page combines direct downloads with internal case briefs so every analytics item remains accessible without broken links.",
  },
];

export const analyticsFeatured: AnalyticsProject[] = [
  {
    title: "Applied Analytics Final Report",
    tag: "Analytics Report",
    summary:
      "A polished long-form analytics report included as a public artifact to show structured analysis, report composition, and executive-facing presentation.",
    outcome:
      "Adds a substantial document to the portfolio that shows Esteban can package analysis into a complete deliverable, not only dashboards and code.",
    stack: ["Applied Analytics", "Reporting", "Business Framing", "Documentation"],
    accent: "var(--cyan)",
    href: "/downloads/analytics-portfolio/reporte-final-analitica-aplicada.pdf",
    hrefLabel: "Open final report",
  },
  {
    title: "Executive BI Dashboard for Commissions, Market, Pareto, and Traders",
    tag: "Power BI",
    summary:
      "A multi-page Power BI project covering commissions, market participation, Pareto logic, decomposition views, and trader-level reporting.",
    outcome:
      "Shows that Esteban can translate data into managerial dashboards, not only backend systems.",
    stack: ["Power BI", "Executive Reporting", "Market Analysis", "Dashboards"],
    accent: "var(--rose)",
    href: "/downloads/analytics-portfolio/proyecto-power-bi-fase-3.pbix",
    hrefLabel: "Download PBIX",
  },
  {
    title: "Property and Price Exploration in Tableau",
    tag: "Data Visualization",
    summary:
      "A Tableau workbook focused on property-level analysis, pricing patterns, housing attributes, and real-estate-oriented visual storytelling.",
    outcome:
      "Expands the public portfolio into analytics communication and exploratory visualization.",
    stack: ["Tableau", "Real Estate", "Visual Analysis", "Storytelling"],
    accent: "var(--cyan)",
    href: "/downloads/analytics-portfolio/proyecto-tableau-fase-3.twb",
    hrefLabel: "Download TWB",
  },
  {
    title: "Fraud Detection Through Graph Analytics",
    tag: "Neo4j",
    summary:
      "A Cypher-based report exploring duplicated cards, shared identity data, overlapping products, and risky loan balances in a fraud scenario.",
    outcome:
      "Introduces graph thinking and pattern analysis without repeating the existing code-lab samples.",
    stack: ["Neo4j", "Cypher", "Fraud Analysis", "Graph Modeling"],
    accent: "var(--amber)",
    href: "/downloads/analytics-portfolio/reporte-fraude-neo4j.pdf",
    hrefLabel: "Open fraud report",
  },
  {
    title: "Healthcare Relationships and Conflict-of-Interest Analysis",
    tag: "Graph Analysis",
    summary:
      "A graph-based healthcare case focused on pharmaceutical influence, overloaded doctors, incompatible treatments, and relational risks.",
    outcome:
      "Broadens the analytical range of the site while staying aligned with structured problem solving.",
    stack: ["Neo4j", "Healthcare", "Risk Analysis", "Relational Data"],
    accent: "var(--rose)",
    href: "/downloads/analytics-portfolio/reporte-medicina-neo4j.pdf",
    hrefLabel: "Open healthcare report",
  },
  {
    title: "Business Intelligence Concept Map",
    tag: "Conceptual Design",
    summary:
      "A visual support piece that complements the applied BI work with a structured overview of business intelligence concepts and relationships.",
    outcome:
      "Shows that Esteban can communicate analytical structure clearly, not only deliver implementation files.",
    stack: ["BI", "Concept Mapping", "Analytical Communication", "Structure"],
    accent: "var(--cyan)",
    href: "/downloads/analytics-portfolio/bi-mapa-conceptual.pdf",
    hrefLabel: "Open concept map",
  },
];

export const analyticsArchive: AnalyticsArchiveItem[] = [
  {
    title: "Big Data Use in a Custom Manufacturing Operation",
    tag: "Big Data",
    summary:
      "Applied essay on operational, customer, logistics, HR, and financial data inside a U.S. manufacturing context.",
    format: "Case Brief",
    href: "/analytics/cases/big-data-uso-tts",
    hrefLabel: "Open case brief",
  },
  {
    title: "Factors Affecting Big Data Adoption in a Service Operations Company",
    tag: "Big Data",
    summary:
      "Technology and organization analysis using TOE and UTAUT perspectives.",
    format: "Case Brief",
    href: "/analytics/cases/big-data-factores-adopcion",
    hrefLabel: "Open case brief",
  },
  {
    title: "Comparative Review of Big Data Technologies",
    tag: "Big Data",
    summary:
      "Comparison across storage, NoSQL, batch, real-time, ML, and visualization platforms.",
    format: "Case Brief",
    href: "/analytics/cases/big-data-cuadro-comparativo",
    hrefLabel: "Open case brief",
  },
  {
    title: "Case Analysis S1",
    tag: "Analytics",
    summary:
      "Case-based analysis of analytics applications in Ecuador with a business-value lens.",
    format: "Case Brief",
    href: "/analytics/cases/analitica-caso-s1",
    hrefLabel: "Open case brief",
  },
  {
    title: "Case Analysis S2 on Public Procurement",
    tag: "Analytics",
    summary:
      "Descriptive and associative analysis of Ecuador's public procurement market using Python.",
    format: "DOCX",
    href: "/downloads/analytics-portfolio/analitica-caso-s2.docx",
  },
  {
    title: "Innovative Analytics Solution S4",
    tag: "Analytics",
    summary:
      "Presentation on cycle prediction, rework risk, and phased implementation planning.",
    format: "PPTX",
    href: "/downloads/analytics-portfolio/analitica-soluciones-innovadoras-s4.pptx",
  },
  {
    title: "Power BI Integrated Project - Phase 1",
    tag: "BI",
    summary:
      "Initial dashboard structure and model setup for the larger reporting workflow.",
    format: "PBIX",
    href: "/downloads/analytics-portfolio/bi-proyecto-fase-1.pbix",
  },
  {
    title: "Power BI Integrated Project - Phase 2",
    tag: "BI",
    summary:
      "Intermediate iteration with more developed layout and expanded analytical views.",
    format: "PBIX",
    href: "/downloads/analytics-portfolio/bi-proyecto-fase-2.pbix",
  },
  {
    title: "Data Visualization Assignment 1",
    tag: "Visualization",
    summary:
      "Additional visual work included to broaden the portfolio footprint beyond one Tableau project.",
    format: "Case Brief",
    href: "/analytics/cases/visualizacion-tarea-1",
    hrefLabel: "Open case brief",
  },
  {
    title: "Graph Database Design and Exploration Task",
    tag: "Graph Databases",
    summary:
      "A Neo4j-oriented task artifact that broadens the graph-analysis footprint of the portfolio.",
    format: "PDF",
    href: "/downloads/analytics-portfolio/analisis-red-tarea-s2.pdf",
  },
  {
    title: "MongoDB Exploration and Querying Task",
    tag: "Databases",
    summary:
      "Collection exploration, filters, counts, and query interpretation in a document database setting.",
    format: "Case Brief",
    href: "/analytics/cases/bases-datos-mongodb",
    hrefLabel: "Open case brief",
  },
  {
    title: "Linear Programming Task S1",
    tag: "Optimization",
    summary:
      "Production-planning model solved with Excel Solver and AMPL.",
    format: "Case Brief",
    href: "/analytics/cases/optimizacion-lineal-s1",
    hrefLabel: "Open case brief",
  },
  {
    title: "Integer Programming Task S3",
    tag: "Optimization",
    summary:
      "Capacity-constrained product-mix optimization with multiple what-if scenarios.",
    format: "Case Brief",
    href: "/analytics/cases/optimizacion-lineal-s3",
    hrefLabel: "Open case brief",
  },
];

export const analyticsCaseStudies: Record<string, AnalyticsCaseStudy> = {
  "big-data-uso-tts": {
    slug: "big-data-uso-tts",
    title: "Big Data Use in a Custom Manufacturing Operation",
    tag: "Big Data",
    summary:
      "A portfolio brief covering how operational, customer, logistics, HR, and finance signals can be combined into a practical big-data agenda for a custom manufacturing business.",
    overview:
      "This case frames big data as an operational leverage layer rather than a buzzword. The work maps where value appears first: demand visibility, rework reduction, delivery coordination, workforce planning, and financial control.",
    whatItShows: [
      "Ability to connect data strategy with real manufacturing pain points.",
      "Clear thinking around operational KPIs, process bottlenecks, and decision support.",
      "Communication that translates technical topics into business language.",
    ],
    focusAreas: ["Operations analytics", "Manufacturing systems", "Decision support", "Data strategy"],
    deliverables: ["Structured case brief", "Business framing", "Portfolio-safe summary artifact"],
  },
  "big-data-factores-adopcion": {
    slug: "big-data-factores-adopcion",
    title: "Factors Affecting Big Data Adoption in a Service Operations Company",
    tag: "Big Data",
    summary:
      "A case brief focused on organizational and technological factors that influence adoption, using TOE and UTAUT-inspired framing.",
    overview:
      "The emphasis here is adoption, not only tooling. The brief examines how readiness, capability, cost, trust, and organizational context shape whether analytics initiatives actually move from idea to execution.",
    whatItShows: [
      "Comfort with analytical frameworks beyond dashboards and coding tasks.",
      "Ability to assess digital transformation risk from both technical and organizational angles.",
      "Experience translating theory into practical business recommendations.",
    ],
    focusAreas: ["Technology adoption", "Organizational readiness", "Analytics strategy", "Change management"],
    deliverables: ["Executive-style brief", "Adoption risk framing", "Portfolio-safe internal page"],
  },
  "big-data-cuadro-comparativo": {
    slug: "big-data-cuadro-comparativo",
    title: "Comparative Review of Big Data Technologies",
    tag: "Big Data",
    summary:
      "A concise comparative brief across storage, NoSQL, batch, real-time, machine learning, and visualization technologies.",
    overview:
      "This piece organizes the ecosystem in a decision-friendly way. Instead of listing tools in isolation, it highlights where each category fits, what tradeoffs matter, and how different layers support different analytical goals.",
    whatItShows: [
      "Technical breadth across databases, processing models, and analytics stacks.",
      "Structured comparison rather than tool-name accumulation.",
      "Ability to communicate architecture choices clearly.",
    ],
    focusAreas: ["Architecture comparison", "Data platforms", "Tool evaluation", "Systems thinking"],
    deliverables: ["Comparative matrix brief", "Architecture notes", "Public-facing case page"],
  },
  "analitica-caso-s1": {
    slug: "analitica-caso-s1",
    title: "Case Analysis S1",
    tag: "Analytics",
    summary:
      "A business-oriented analytics case brief focused on applying descriptive and evaluative analysis to real decision scenarios.",
    overview:
      "The goal of this piece is to show analytical framing discipline: understanding the problem, identifying the right variables, and turning observations into usable recommendations instead of generic commentary.",
    whatItShows: [
      "Ability to structure analytical work around decision quality.",
      "Clear business framing and concise recommendation writing.",
      "A broader analytical footprint beyond code-heavy portfolio pieces.",
    ],
    focusAreas: ["Analytical reasoning", "Business framing", "Decision support", "Structured reporting"],
    deliverables: ["Case summary", "Recommendation-focused writeup", "Internal brief page"],
  },
  "visualizacion-tarea-1": {
    slug: "visualizacion-tarea-1",
    title: "Data Visualization Assignment 1",
    tag: "Visualization",
    summary:
      "A visualization-focused case brief showing additional exploratory and presentation work beyond the main Tableau artifact.",
    overview:
      "This entry exists to widen the visual layer of the analytics portfolio. It highlights the ability to shape a narrative through layout, comparison, and explanatory structure rather than relying only on one flagship dashboard.",
    whatItShows: [
      "Additional proof of visual communication skills.",
      "Comfort with presenting data stories to non-technical audiences.",
      "Range beyond operational systems and ETL work.",
    ],
    focusAreas: ["Data storytelling", "Visualization design", "Analytical communication", "Presentation clarity"],
    deliverables: ["Visualization brief", "Portfolio narrative support", "Viewable internal artifact"],
  },
  "bases-datos-mongodb": {
    slug: "bases-datos-mongodb",
    title: "MongoDB Exploration and Querying Task",
    tag: "Databases",
    summary:
      "A case brief centered on collection design, document queries, filtering patterns, counts, and interpretation in a MongoDB context.",
    overview:
      "The portfolio value here is database range. This brief demonstrates familiarity with document-oriented thinking, query construction, and the differences between relational and NoSQL-style exploration workflows.",
    whatItShows: [
      "Range across both relational and document databases.",
      "Understanding of exploratory querying and data inspection patterns.",
      "Ability to communicate technical database work in a digestible format.",
    ],
    focusAreas: ["MongoDB", "NoSQL", "Query logic", "Database literacy"],
    deliverables: ["Database case brief", "Query-focused notes", "Internal page instead of missing raw file"],
  },
  "optimizacion-lineal-s1": {
    slug: "optimizacion-lineal-s1",
    title: "Linear Programming Task S1",
    tag: "Optimization",
    summary:
      "A case brief showing production-planning logic solved with optimization tooling such as Excel Solver and AMPL.",
    overview:
      "This work adds operations research signal to the portfolio. It focuses on objective functions, constraints, and structured decision tradeoffs rather than only descriptive analytics or dashboards.",
    whatItShows: [
      "Comfort with optimization and constrained decision modeling.",
      "Ability to bridge engineering thinking with business tradeoffs.",
      "Analytical depth that complements ETL and BI work.",
    ],
    focusAreas: ["Linear programming", "Operations research", "Constraint modeling", "Decision optimization"],
    deliverables: ["Optimization brief", "Modeling summary", "Public-safe internal case page"],
  },
  "optimizacion-lineal-s3": {
    slug: "optimizacion-lineal-s3",
    title: "Integer Programming Task S3",
    tag: "Optimization",
    summary:
      "A case brief focused on capacity-constrained product mix optimization and scenario-based decision making.",
    overview:
      "This entry pushes the optimization layer a bit further by making discrete choices matter. It highlights scenario evaluation, capacity planning, and the impact of operational limits on final recommendations.",
    whatItShows: [
      "Capability in integer programming and what-if analysis.",
      "Structured reasoning under operational constraints.",
      "A stronger quantitative layer inside the analytics page.",
    ],
    focusAreas: ["Integer programming", "Scenario analysis", "Capacity planning", "Optimization"],
    deliverables: ["Scenario brief", "Optimization framing", "Internal portfolio case page"],
  },
};
