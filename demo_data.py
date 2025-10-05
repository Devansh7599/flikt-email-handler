#!/usr/bin/env python3
"""
demo_data.py

Generates and serves a large, realistic demo email dataset for the Email Filter
and Dashboard application. Data is synthesized deterministically for repeatable
results and covers a broad date range so UI filtering can be exercised.
"""

from __future__ import annotations

from datetime import datetime, timedelta
from typing import List, Dict
import random


_CACHED_DATASET: List[Dict] = []


def _generate_large_dataset() -> List[Dict]:
    """Create a large, realistic dataset (~3,000 emails) spanning 180 days.

    Returns a list of dicts with keys: name, email, subject, body, date.
    """
    senders = [
        ("Himanshu", "himanshu@example.com"),
        ("Sakher", "sakher@company.com"),
        ("Mayank", "mayank@techcorp.com"),
        ("sandhya", "sandhya@startup.io"),
        ("vinod", "vinod@consulting.com"),
        ("Priya Nair", "priya.nair@acme.co"),
        ("aviral", "aviral@supplychain.io"),
        ("Aisha Khan", "aisha.khan@fintech.app"),
        ("Nora", "nora@retailhub.com"),
        ("Lucky", "lucky@hardware.cn"),
        ("shobhit", "shobhit@healthcare.org"),
        ("chavi", "chavi@edutech.edu"),
    ]

    subjects = [
        "Meeting Reminder",
        "Project Update",
        "Invoice Attached",
        "Release Notes",
        "Action Required",
        "Budget Approval",
        "Welcome Aboard",
        "Weekly Report",
        "Customer Feedback",
        "Outage Postmortem",
        "Contract Review",
        "Security Notice",
    ]

    body_snippets = [
        "Please find the details in the attached document. Let me know if you have questions.",
        "We are on track against the current milestones and expect to hit the deadline.",
        "This is a reminder for the meeting scheduled tomorrow at 2 PM.",
        "The latest build includes performance improvements and bug fixes across modules.",
        "Kindly review and approve at your earliest convenience.",
        "Thanks for your prompt attention to this matter.",
        "Summarizing this week’s progress and next steps for the team.",
        "Please review the notes and provide your feedback by EOD.",
        "We observed an increase in engagement week over week.",
        "Action items are listed at the end of this message.",
    ]

    # Seed for deterministic output
    random.seed(12345)

    today = datetime.today().replace(hour=0, minute=0, second=0, microsecond=0)
    start_range = today - timedelta(days=179)
    end_range = today

    dataset: List[Dict] = []
    current = start_range
    while current <= end_range:
        emails_today = random.randint(5, 30)  # denser data
        for _ in range(emails_today):
            sender_name, sender_email = random.choice(senders)
            subject = f"{random.choice(subjects)} - {current.strftime('%b %d, %Y')}"
            body = (
                random.choice(body_snippets) + " " +
                random.choice(body_snippets) + " " +
                random.choice(body_snippets)
            )
            # Random time during the day
            dt = current + timedelta(
                hours=random.randint(0, 23),
                minutes=random.randint(0, 59),
                seconds=random.randint(0, 59),
            )
            dataset.append({
                'name': sender_name,
                'email': sender_email,
                'subject': subject,
                'body': body,
                'date': dt,
            })
        current += timedelta(days=1)

    # Sort by date descending to show latest first
    dataset.sort(key=lambda x: x['date'], reverse=True)
    return dataset


def get_demo_dataset() -> List[Dict]:
    """Return the cached large dataset, generating it on first use."""
    global _CACHED_DATASET
    if not _CACHED_DATASET:
        _CACHED_DATASET = _generate_large_dataset()
    return _CACHED_DATASET


def load_demo_emails_between(start_date: datetime, end_date: datetime) -> List[Dict]:
    """Filter the large dataset between start_date and end_date (inclusive)."""
    data = get_demo_dataset()
    start_key = start_date.replace(hour=0, minute=0, second=0, microsecond=0)
    end_key = end_date.replace(hour=23, minute=59, second=59, microsecond=0)
    return [item for item in data if start_key <= item['date'] <= end_key]


