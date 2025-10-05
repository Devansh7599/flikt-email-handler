# Email Filter & Dashboard

A desktop application built with Python and Tkinter to fetch, filter, and visualize emails from any IMAP-supported email account.

## Features

- **Multiple Providers**: Connect to Gmail, Outlook, or a custom IMAP server.
- **Date Filtering**: Select a precise date range to fetch emails.
- **Demo Mode**: Test the application's features safely with built-in demo data, no credentials required.
- **Interactive Dashboard**: View emails in a sortable, searchable table.
- **Data Export**: Export filtered email data to CSV or JSON formats.
- **Modern UI**: Includes light and dark themes, tooltips, and a notification system.
- **Real-time Search**: Instantly filter emails in the dashboard by sender, subject, or body content.

## Prerequisites

- Python 3.7 or higher

## Installation

This application uses only Python's standard libraries, so no external packages are required.

1.  Clone or download the project files.
2.  Ensure you have Python 3.7+ installed.

## How to Run

Open your terminal or command prompt, navigate to the project directory, and run:

```bash
python email_filter_dashboard.py
```

## Usage Guide

### 1. Using Demo Mode (Recommended for First Use)

1.  Launch the application.
2.  Keep the **"Use Demo Data (for testing)"** checkbox checked.
3.  Select a start and end date.
4.  Click **Fetch Emails**.
5.  Once fetching is complete, click **Open Dashboard** to view the results.

### 2. Connecting to a Real Email Account

1.  Uncheck the "Use Demo Data" box.
2.  Select your email provider (e.g., Gmail, Outlook) from the dropdown.
3.  Enter your email address and password.
    -   **Gmail Users**: You must enable IMAP in your Gmail settings and generate an **App Password**. Use this App Password in the password field, not your regular account password.
4.  Select the desired date range.
5.  Click **Fetch Emails**.
6.  Click **Open Dashboard** to view your emails.