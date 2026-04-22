from datetime import datetime
import re


def convert_to_iso_string(value):
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


def extract_area_from_notes(notes):
    if not isinstance(notes, str):
        return "Not specified"
    match = re.search(
        r"(\d+\.?\d*)\s*(?:sq(?:uare)?\s?ft|sft|sf|ft|area|s?f|sq)",
        notes,
        re.IGNORECASE,
    )
    return match.group(1) if match else "Not specified"


def build_activity_payload(job_activities):
    activity_data = {
        "Template": {"notes": None},
        "CAD": {"status": None, "assigned_to": None, "date": None},
        "Fab": {"status": None, "date": None},
    }

    for activity in job_activities:
        activity_name = activity.JobActivityTypeName
        if activity_name not in activity_data:
            continue

        if "status" in activity_data[activity_name]:
            activity_data[activity_name]["status"] = str(activity.JobActivityStatusName)

        if "assigned_to" in activity_data[activity_name]:
            assignees = []
            enumerator = activity.Assignees.GetEnumerator()
            while enumerator.MoveNext():
                assignees.append(str(enumerator.Current.AssigneeName))
            activity_data[activity_name]["assigned_to"] = ", ".join(assignees)

        if "date" in activity_data[activity_name]:
            activity_data[activity_name]["date"] = convert_to_iso_string(activity.StartDate)

        if activity_name == "Template":
            activity_data["Template"]["notes"] = str(activity.Notes or "").strip() or "No notes"

    return {
        "template_notes": activity_data["Template"]["notes"],
        "template_area": extract_area_from_notes(activity_data["Template"]["notes"]),
        "cad_status": activity_data["CAD"]["status"],
        "cad_assigned_to": activity_data["CAD"]["assigned_to"],
        "cad_date": activity_data["CAD"]["date"],
        "fab_status": activity_data["Fab"]["status"],
        "fab_date": activity_data["Fab"]["date"],
    }
