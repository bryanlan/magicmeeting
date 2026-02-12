#!/usr/bin/env python3
"""
Parse Outlook calendar JSON responses and format them efficiently.
Handles the large JSON responses from the MCP Outlook server.

Usage:
    python parse_calendar.py <json_file> [--date MMDDYYYY] [--range START END] [--format summary|detailed]

Examples:
    python parse_calendar.py calendar.txt --date 01302026
    python parse_calendar.py calendar.txt --range 01302026 02052026
    python parse_calendar.py calendar.txt --date 01302026 --format detailed
"""

import json
import sys
import argparse
from datetime import datetime
from typing import Optional


def parse_date(date_str: str) -> Optional[datetime]:
    """Parse date string in M/D/YYYY H:MM AM/PM format."""
    try:
        return datetime.strptime(date_str, "%m/%d/%Y %I:%M %p")
    except ValueError:
        try:
            return datetime.strptime(date_str, "%m/%d/%Y")
        except ValueError:
            return None


def filter_events_by_date(events: list, target_date: str) -> list:
    """Filter events for a specific date (MMDDYYYY format)."""
    # Convert MMDDYYYY to M/D/YYYY pattern for matching
    month = target_date[0:2].lstrip('0') or '0'
    day = target_date[2:4].lstrip('0') or '0'
    year = target_date[4:8]
    date_pattern = f"{month}/{day}/{year}"

    return [e for e in events if date_pattern in e.get('start', '')]


def filter_events_by_range(events: list, start_date: str, end_date: str) -> list:
    """Filter events within a date range (MMDDYYYY format)."""
    start = datetime.strptime(start_date, "%m%d%Y")
    end = datetime.strptime(end_date, "%m%d%Y").replace(hour=23, minute=59, second=59)

    filtered = []
    for e in events:
        event_start = parse_date(e.get('start', ''))
        if event_start and start <= event_start <= end:
            filtered.append(e)
    return filtered


def format_time(dt_str: str) -> str:
    """Extract just the time portion from a datetime string."""
    dt = parse_date(dt_str)
    if dt:
        return dt.strftime("%I:%M %p").lstrip('0')
    return dt_str


def format_event_summary(event: dict) -> str:
    """Format a single event as a summary line."""
    start = format_time(event.get('start', ''))
    end = format_time(event.get('end', ''))
    subject = event.get('subject', 'No subject')
    status = event.get('busyStatus', 'Unknown')
    location = event.get('location', '')

    status_indicator = {
        'Free': '[Free]',
        'Busy': '[Busy]',
        'Tentative': '[Tentative]',
        'Out of Office': '[OOF]'
    }.get(status, f'[{status}]')

    line = f"{start} - {end}: {subject} {status_indicator}"
    if location and len(location) < 50:
        line += f" @ {location}"
    return line


def format_event_detailed(event: dict) -> str:
    """Format a single event with full details."""
    lines = []
    lines.append(f"  Subject: {event.get('subject', 'No subject')}")
    lines.append(f"  Time: {event.get('start', '')} - {event.get('end', '')}")
    lines.append(f"  Status: {event.get('busyStatus', 'Unknown')}")
    if event.get('location'):
        lines.append(f"  Location: {event.get('location')}")
    if event.get('organizer'):
        lines.append(f"  Organizer: {event.get('organizer')}")
    if event.get('attendees'):
        attendee_names = [a.get('name', '') for a in event['attendees'][:5]]
        if len(event['attendees']) > 5:
            attendee_names.append(f"... +{len(event['attendees']) - 5} more")
        lines.append(f"  Attendees: {', '.join(attendee_names)}")
    if event.get('isRecurring'):
        lines.append("  [Recurring]")
    return '\n'.join(lines)


def group_by_time_of_day(events: list) -> dict:
    """Group events into morning, afternoon, evening."""
    groups = {'All Day': [], 'Morning': [], 'Afternoon': [], 'Evening': []}

    for event in events:
        start_dt = parse_date(event.get('start', ''))
        end_dt = parse_date(event.get('end', ''))

        # Check for all-day events (midnight to midnight or 24+ hours)
        if start_dt and end_dt:
            duration = (end_dt - start_dt).total_seconds() / 3600
            if duration >= 23 or (start_dt.hour == 0 and start_dt.minute == 0):
                groups['All Day'].append(event)
                continue

        if start_dt:
            hour = start_dt.hour
            if hour < 12:
                groups['Morning'].append(event)
            elif hour < 17:
                groups['Afternoon'].append(event)
            else:
                groups['Evening'].append(event)
        else:
            groups['Morning'].append(event)

    return groups


def load_mcp_response(file_path: str) -> list:
    """Load and parse MCP Outlook response JSON."""
    with open(file_path, 'r', encoding='utf-8') as f:
        data = json.load(f)

    # MCP responses are wrapped in [{type: "text", text: "<json>"}]
    if isinstance(data, list) and len(data) > 0 and 'text' in data[0]:
        return json.loads(data[0]['text'])
    return data


def main():
    parser = argparse.ArgumentParser(description='Parse Outlook calendar JSON')
    parser.add_argument('json_file', help='Path to the JSON file')
    parser.add_argument('--date', help='Filter by date (MMDDYYYY)')
    parser.add_argument('--range', nargs=2, metavar=('START', 'END'),
                        help='Filter by date range (MMDDYYYY MMDDYYYY)')
    parser.add_argument('--format', choices=['summary', 'detailed'],
                        default='summary', help='Output format')
    parser.add_argument('--json', action='store_true',
                        help='Output as JSON instead of formatted text')

    args = parser.parse_args()

    # Load events
    events = load_mcp_response(args.json_file)

    # Filter by date or range
    if args.date:
        events = filter_events_by_date(events, args.date)
    elif args.range:
        events = filter_events_by_range(events, args.range[0], args.range[1])

    # Sort by start time
    events.sort(key=lambda e: parse_date(e.get('start', '')) or datetime.min)

    if args.json:
        # Output filtered events as JSON (without body field to reduce size)
        clean_events = []
        for e in events:
            clean = {k: v for k, v in e.items() if k != 'body'}
            clean_events.append(clean)
        print(json.dumps(clean_events, indent=2))
        return

    # Group and format
    groups = group_by_time_of_day(events)

    for period, period_events in groups.items():
        if not period_events:
            continue
        print(f"\n**{period}:**")
        for event in period_events:
            if args.format == 'detailed':
                print(format_event_detailed(event))
                print()
            else:
                print(f"- {format_event_summary(event)}")


if __name__ == '__main__':
    main()
