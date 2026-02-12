#!/usr/bin/env python3
"""
parse_calendar.py - Parse and filter Outlook calendar JSON exports.

Usage:
    python parse_calendar.py <file> --date MMDDYYYY [--format summary|detailed] [--json]
    python parse_calendar.py <file> --range MMDDYYYY MMDDYYYY [--format summary|detailed] [--json]

Output format groups events by time of day (Morning/Afternoon/Evening).
"""

import argparse
import json
import sys
from datetime import datetime
from typing import Any


def parse_date_arg(date_str: str) -> datetime:
    """Parse MMDDYYYY format to datetime."""
    return datetime.strptime(date_str, "%m%d%Y")


def parse_event_datetime(dt_str: str) -> datetime:
    """Parse event datetime string like '2/12/2026 08:00 AM'."""
    # Handle various formats from Outlook
    for fmt in ["%m/%d/%Y %I:%M %p", "%m/%d/%Y %H:%M"]:
        try:
            return datetime.strptime(dt_str, fmt)
        except ValueError:
            continue
    raise ValueError(f"Cannot parse datetime: {dt_str}")


def get_time_of_day(dt: datetime) -> str:
    """Categorize time into Morning/Afternoon/Evening."""
    hour = dt.hour
    if hour < 12:
        return "Morning"
    elif hour < 17:
        return "Afternoon"
    else:
        return "Evening"


def format_time(dt: datetime) -> str:
    """Format datetime to time string like '9:00 AM'."""
    return dt.strftime("%I:%M %p").lstrip("0")


def get_status_indicator(busy_status: str) -> str:
    """Map busyStatus to display indicator."""
    status_map = {
        "Busy": "[Busy]",
        "Tentative": "[Tentative]",
        "Free": "[Free]",
        "Out of Office": "[OOF]",
        "WorkingElsewhere": "[WFH]",
    }
    return status_map.get(busy_status, f"[{busy_status}]")


def is_all_day_event(event: dict) -> bool:
    """Check if event is an all-day event (starts at midnight, spans full day)."""
    try:
        start = parse_event_datetime(event["start"])
        end = parse_event_datetime(event["end"])
        return start.hour == 0 and start.minute == 0 and (end - start).days >= 1
    except (ValueError, KeyError):
        return False


def event_in_date_range(event: dict, start_date: datetime, end_date: datetime) -> bool:
    """Check if event falls within date range."""
    try:
        event_start = parse_event_datetime(event["start"])
        event_date = event_start.date()
        return start_date.date() <= event_date <= end_date.date()
    except (ValueError, KeyError):
        return False


def load_calendar_json(filepath: str) -> list[dict]:
    """Load and extract events from calendar JSON file."""
    with open(filepath, "r", encoding="utf-8") as f:
        data = json.load(f)

    # Handle the MCP response format: array with text blocks
    events = []
    for item in data:
        if isinstance(item, dict) and item.get("type") == "text":
            text = item.get("text", "")
            # Skip date context messages
            if text.startswith("{"):
                try:
                    parsed = json.loads(text)
                    if "events" in parsed:
                        events.extend(parsed["events"])
                except json.JSONDecodeError:
                    continue

    return events


def format_event_summary(event: dict) -> str:
    """Format event in summary format: TIME - TIME: Subject [Status] @ Location"""
    try:
        start = parse_event_datetime(event["start"])
        end = parse_event_datetime(event["end"])
        start_time = format_time(start)
        end_time = format_time(end)
    except ValueError:
        start_time = event.get("start", "?")
        end_time = event.get("end", "?")

    subject = event.get("subject", "(No subject)")
    status = get_status_indicator(event.get("busyStatus", "Busy"))
    location = event.get("location", "")

    line = f"- {start_time} - {end_time}: {subject} {status}"
    if location and location.strip():
        line += f" @ {location}"

    return line


def format_event_detailed(event: dict) -> str:
    """Format event with additional details."""
    lines = [format_event_summary(event)]

    organizer = event.get("organizer", "")
    if organizer:
        lines.append(f"    Organizer: {organizer}")

    attendees = event.get("attendees", [])
    if attendees:
        attendee_names = [a.get("name", a.get("email", "?")) for a in attendees[:5]]
        if len(attendees) > 5:
            attendee_names.append(f"... +{len(attendees) - 5} more")
        lines.append(f"    Attendees: {', '.join(attendee_names)}")

    if event.get("isRecurring"):
        lines.append("    (Recurring)")

    return "\n".join(lines)


def format_output(events: list[dict], format_type: str, target_date: datetime = None) -> str:
    """Format events grouped by time of day."""
    if not events:
        date_str = target_date.strftime("%m/%d/%Y") if target_date else "specified range"
        return f"No events found for {date_str}"

    # Separate all-day events from timed events
    all_day = []
    timed = []
    for event in events:
        if is_all_day_event(event):
            all_day.append(event)
        else:
            timed.append(event)

    # Sort timed events by start time
    timed.sort(key=lambda e: parse_event_datetime(e["start"]))

    # Group by time of day
    groups = {"All Day": all_day, "Morning": [], "Afternoon": [], "Evening": []}
    for event in timed:
        try:
            start = parse_event_datetime(event["start"])
            tod = get_time_of_day(start)
            groups[tod].append(event)
        except ValueError:
            groups["Morning"].append(event)  # Default fallback

    # Build output
    output_lines = []
    formatter = format_event_detailed if format_type == "detailed" else format_event_summary

    for period in ["All Day", "Morning", "Afternoon", "Evening"]:
        if groups[period]:
            output_lines.append(f"\n**{period}:**")
            for event in groups[period]:
                output_lines.append(formatter(event))

    return "\n".join(output_lines).strip()


def main():
    parser = argparse.ArgumentParser(description="Parse Outlook calendar JSON exports")
    parser.add_argument("file", help="Path to calendar JSON file")
    parser.add_argument("--date", metavar="MMDDYYYY", help="Filter by single date")
    parser.add_argument("--range", nargs=2, metavar=("START", "END"), help="Filter by date range (MMDDYYYY MMDDYYYY)")
    parser.add_argument("--format", choices=["summary", "detailed"], default="summary", help="Output format")
    parser.add_argument("--json", action="store_true", help="Output as JSON")

    args = parser.parse_args()

    if not args.date and not args.range:
        parser.error("Either --date or --range is required")

    # Parse date filters
    if args.date:
        start_date = parse_date_arg(args.date)
        end_date = start_date
    else:
        start_date = parse_date_arg(args.range[0])
        end_date = parse_date_arg(args.range[1])

    # Load and filter events
    try:
        events = load_calendar_json(args.file)
    except FileNotFoundError:
        print(f"Error: File not found: {args.file}", file=sys.stderr)
        sys.exit(1)
    except json.JSONDecodeError as e:
        print(f"Error: Invalid JSON: {e}", file=sys.stderr)
        sys.exit(1)

    filtered = [e for e in events if event_in_date_range(e, start_date, end_date)]

    # Output
    if args.json:
        print(json.dumps(filtered, indent=2))
    else:
        print(format_output(filtered, args.format, start_date if args.date else None))


if __name__ == "__main__":
    main()
