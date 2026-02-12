#!/usr/bin/env python3
"""
Find optimal meeting times across multiple calendars.
Analyzes availability to find slots where the most people are free.

Usage:
    python find_meeting_times.py <calendars_json> --duration 60 --top 5 --my-calendar <my_calendar_json>

Input format: JSON file with structure:
{
    "Person Name": [list of events],
    ...
}

Output: Top N time slots with availability analysis.
"""

import json
import argparse
from datetime import datetime, timedelta
from collections import defaultdict
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


def get_busy_periods(events: list) -> list:
    """Extract busy/tentative periods from events."""
    periods = []
    for event in events:
        status = event.get('busyStatus', 'Busy')
        if status in ('Free',):
            continue
        start = parse_date(event.get('start', ''))
        end = parse_date(event.get('end', ''))
        if start and end:
            periods.append({
                'start': start,
                'end': end,
                'status': status,
                'subject': event.get('subject', 'Busy')
            })
    return periods


def is_person_available(periods: list, slot_start: datetime, slot_end: datetime) -> tuple:
    """
    Check if person is available during a time slot.
    Returns (availability_status, conflicting_event_or_none)
    Status: 'free', 'tentative', 'busy', 'ooo'
    """
    for period in periods:
        # Check for overlap
        if period['start'] < slot_end and period['end'] > slot_start:
            status = period['status'].lower()
            if 'out' in status or 'oof' in status.lower():
                return ('ooo', period)
            elif 'tentative' in status:
                return ('tentative', period)
            else:
                return ('busy', period)
    return ('free', None)


def generate_time_slots(start_date: str, end_date: str, duration_minutes: int,
                        work_start: int = 9, work_end: int = 17) -> list:
    """Generate possible meeting slots within work hours."""
    start = datetime.strptime(start_date, "%m%d%Y")
    end = datetime.strptime(end_date, "%m%d%Y")
    slots = []

    current = start.replace(hour=work_start, minute=0)
    while current.date() <= end.date():
        # Skip weekends
        if current.weekday() < 5:  # Monday = 0, Friday = 4
            slot_end = current + timedelta(minutes=duration_minutes)
            if slot_end.hour < work_end or (slot_end.hour == work_end and slot_end.minute == 0):
                slots.append((current, slot_end))
            # Move to next slot (30-minute increments)
            current += timedelta(minutes=30)
            # Check if we've gone past work hours
            if current.hour >= work_end:
                # Move to next day
                current = (current + timedelta(days=1)).replace(hour=work_start, minute=0)
        else:
            # Skip to Monday
            days_until_monday = 7 - current.weekday()
            current = (current + timedelta(days=days_until_monday)).replace(hour=work_start, minute=0)

    return slots


def analyze_slot(slot_start: datetime, slot_end: datetime,
                 people_busy_periods: dict, my_busy_periods: list = None) -> dict:
    """Analyze a time slot for all people."""
    result = {
        'start': slot_start,
        'end': slot_end,
        'free': [],
        'tentative': [],
        'busy': [],
        'ooo': [],
        'my_conflicts': []
    }

    for person, periods in people_busy_periods.items():
        status, conflict = is_person_available(periods, slot_start, slot_end)
        if status == 'free':
            result['free'].append(person)
        elif status == 'tentative':
            result['tentative'].append({'name': person, 'event': conflict.get('subject') if conflict else None})
        elif status == 'ooo':
            result['ooo'].append(person)
        else:
            result['busy'].append({'name': person, 'event': conflict.get('subject') if conflict else None})

    # Check my calendar
    if my_busy_periods:
        for period in my_busy_periods:
            if period['start'] < slot_end and period['end'] > slot_start:
                result['my_conflicts'].append({
                    'subject': period.get('subject', 'Busy'),
                    'status': period.get('status', 'Busy')
                })

    # Score: prioritize free, then tentative (tentative counts as 0.5)
    result['available_count'] = len(result['free']) + len(result['tentative']) * 0.5
    result['total_people'] = len(people_busy_periods)

    return result


def format_slot_result(result: dict, show_my_calendar: bool = True) -> str:
    """Format a single slot result for display."""
    lines = []
    start = result['start'].strftime("%a %m/%d %I:%M %p")
    end = result['end'].strftime("%I:%M %p")
    available = int(result['available_count'])
    total = result['total_people']

    lines.append(f"**{start} - {end}** ({available}/{total} available)")

    if result['free']:
        lines.append(f"  Free: {', '.join(result['free'])}")

    if result['tentative']:
        tentative_str = ', '.join([f"{t['name']} ({t['event']})" for t in result['tentative']])
        lines.append(f"  Tentative: {tentative_str}")

    if result['busy']:
        busy_str = ', '.join([f"{b['name']}" for b in result['busy']])
        lines.append(f"  Can't make it: {busy_str}")

    if result['ooo']:
        lines.append(f"  Out of Office: {', '.join(result['ooo'])}")

    if show_my_calendar and result['my_conflicts']:
        conflicts = ', '.join([f"{c['subject']} [{c['status']}]" for c in result['my_conflicts']])
        lines.append(f"  Your conflicts: {conflicts}")

    return '\n'.join(lines)


def main():
    parser = argparse.ArgumentParser(description='Find optimal meeting times')
    parser.add_argument('calendars_json', help='JSON file with all calendars')
    parser.add_argument('--duration', type=int, default=60, help='Meeting duration in minutes')
    parser.add_argument('--top', type=int, default=5, help='Number of top slots to show')
    parser.add_argument('--start', required=True, help='Start date (MMDDYYYY)')
    parser.add_argument('--end', required=True, help='End date (MMDDYYYY)')
    parser.add_argument('--my-calendar', help='Your calendar JSON file')
    parser.add_argument('--work-start', type=int, default=9, help='Work day start hour (0-23)')
    parser.add_argument('--work-end', type=int, default=17, help='Work day end hour (0-23)')
    parser.add_argument('--json', action='store_true', help='Output as JSON')

    args = parser.parse_args()

    # Load calendars
    with open(args.calendars_json, 'r', encoding='utf-8') as f:
        calendars = json.load(f)

    # Extract busy periods for each person
    people_busy_periods = {}
    for person, events in calendars.items():
        people_busy_periods[person] = get_busy_periods(events)

    # Load my calendar if provided
    my_busy_periods = None
    if args.my_calendar:
        with open(args.my_calendar, 'r', encoding='utf-8') as f:
            my_data = json.load(f)
            # Handle MCP response format
            if isinstance(my_data, list) and len(my_data) > 0 and 'text' in my_data[0]:
                my_events = json.loads(my_data[0]['text'])
            else:
                my_events = my_data
            my_busy_periods = get_busy_periods(my_events)

    # Generate time slots
    slots = generate_time_slots(args.start, args.end, args.duration,
                                 args.work_start, args.work_end)

    # Analyze each slot
    results = []
    for slot_start, slot_end in slots:
        result = analyze_slot(slot_start, slot_end, people_busy_periods, my_busy_periods)
        results.append(result)

    # Sort by availability (highest first), then by date
    results.sort(key=lambda r: (-r['available_count'], r['start']))

    # Take top N
    top_results = results[:args.top]

    if args.json:
        # Convert to JSON-serializable format
        json_results = []
        for r in top_results:
            json_results.append({
                'start': r['start'].isoformat(),
                'end': r['end'].isoformat(),
                'available_count': r['available_count'],
                'total_people': r['total_people'],
                'free': r['free'],
                'tentative': [t['name'] for t in r['tentative']],
                'busy': [b['name'] for b in r['busy']],
                'ooo': r['ooo'],
                'my_conflicts': r['my_conflicts']
            })
        print(json.dumps(json_results, indent=2))
    else:
        print(f"\nTop {len(top_results)} meeting slots for {args.duration}-minute meeting:\n")
        for i, result in enumerate(top_results, 1):
            print(f"{i}. {format_slot_result(result, my_busy_periods is not None)}")
            print()


if __name__ == '__main__':
    main()
